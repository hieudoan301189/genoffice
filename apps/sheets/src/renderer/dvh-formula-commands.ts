// DVH functions whose job is a workbook change: defined names, notes, print
// setup, row heights and rich-text formatting. Ported from DVH-Tool/Function/
// clsFun_Name.cs (AddName, DelName, ShowName), clsFun_Comment.cs (TextCom,
// DelCom), clsFun_Print.cs (PrintArea, PrintTitle), clsFun_FitRow.cs,
// clsFun_JoinFormat.cs and clsFun_FindFormat.cs.
//
// Excel-DNA ran these changes from inside the recalculation (macro-type
// functions and QueueAsMacro). Here a function only computes what it would do
// and returns its status text; the change itself is queued as a "formula
// command" keyed by the formula cell and a fingerprint of the request, and
// runs once after the calculation ends, inside one undo step:
//  - a formula the user entered or edited this session runs on its first
//    calculation; formulas loaded from a file are only remembered on the
//    workbook's first calculation (the file already carries their effect);
//  - afterwards a command reruns only when its fingerprint changes (JoinFormat
//    and FindFormat fingerprint their source text, so they follow edits);
//  - ClearFindCache/ClearFormatCache forget the fingerprints, as clearing the
//    add-in's caches made it re-apply the formats.
import {
  BaseFunction,
  FunctionType,
  IFunctionService,
  StringValueObject,
  type BaseReferenceObject,
  type BaseValueObject,
  type IFunctionInfo,
} from '@univerjs/engine-formula'
import {
  ICommandService,
  IUniverInstanceService,
  ObjectMatrix,
  type ICellData,
  type IObjectMatrixPrimitiveType,
  type IStyleData,
  type Nullable,
} from '@univerjs/core'
import { SetRangeValuesMutation } from '@univerjs/sheets'
import { IDescriptionService } from '@univerjs/sheets-formula'
import type { FWorksheet } from '@univerjs/sheets/facade'
import { isDvhDefinedName } from '@genoffice/dvh-model'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookRichRun } from '@genoffice/xlsx-gateway/shared/edit-schemas'

import {
  asReference,
  boundsOf,
  isBlankArg,
  isFalseLike,
  matrixOf,
  scalarOf,
} from './dvh-format-functions'
import { registerDvhAliases } from './dvh-function-aliases'
import { fitRowHeights } from './dvh-home-actions'
import { nameMatchesCategory, workbookNames } from './dvh-workbook-functions'
import { extractRichText } from './edit-journal'
import { beginUndoBatch } from './op-executor'
import { toRichTextDocument } from './univer-sync'
import type { UniverRuntime } from './univer-state'

type Scalar = string | number | boolean

interface Bounds {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

const MAX_COMMAND_CELLS = 50_000

// ---------------- the command queue ----------------

/** What the app lends the queue: journal lookups, page setup and messages. */
export interface DvhCommandHost {
  /** True when the user entered or edited this cell's formula this session. */
  isUserFormula(sheetId: string, row: number, column: number): boolean
  /** True once the workbook is fully loaded (streamed files load in steps). */
  isLoaded(): boolean
  /** Records print area / print titles the way Page Layout does; false without a file. */
  recordPrintSetup(
    sheetId: string,
    patch: { printArea?: string | null; printTitles?: string | null },
  ): boolean
  notify(message: string, kind: 'info' | 'error'): void
  /** Places a picture anchored at a cell under `name`, replacing an earlier one of that name. */
  placePicture(
    sheetId: string,
    row: number,
    column: number,
    name: string,
    image: { dataUrl: string; mediaType: string; width: number; height: number },
  ): boolean
}

export interface CommandRequest {
  readonly unitId: string
  readonly sheetId: string
  readonly row: number
  readonly column: number
  /** upper-case function name */
  readonly fn: string
  /** fingerprint: the same request twice runs once */
  readonly key: string
  /** returns a status-bar message (or nothing); async runs land after the undo step */
  readonly run: () => string | void | Promise<string | void>
}

interface PendingRequest extends CommandRequest {
  /** the workbook had finished its first calculation when the request came */
  readonly afterFirstCalculation: boolean
}

const ownerKey = (r: Pick<CommandRequest, 'unitId' | 'sheetId' | 'row' | 'column'>): string =>
  `${r.unitId}\u0000${r.sheetId}\u0000${r.row}\u0000${r.column}`

export class DvhCommandQueue {
  private readonly done = new Map<string, { fn: string; key: string }>()
  private readonly pending = new Map<string, PendingRequest>()
  private readonly calculated = new Set<string>()

  request(request: CommandRequest): void {
    this.pending.set(ownerKey(request), {
      ...request,
      afterFirstCalculation: this.calculated.has(request.unitId),
    })
  }

  /**
   * Decides which pending requests run: a new fingerprint runs, an unchanged
   * one does not; a first sighting from the file's own first calculation is
   * only remembered.
   */
  take(isUserFormula: DvhCommandHost['isUserFormula']): PendingRequest[] {
    const runs: PendingRequest[] = []
    for (const [key, request] of this.pending) {
      const previous = this.done.get(key)
      this.done.set(key, { fn: request.fn, key: request.key })
      if (previous && previous.fn === request.fn && previous.key === request.key) continue
      const firstSighting = !previous || previous.fn !== request.fn
      if (
        firstSighting &&
        !request.afterFirstCalculation &&
        !isUserFormula(request.sheetId, request.row, request.column)
      )
        continue
      runs.push(request)
    }
    this.pending.clear()
    return runs
  }

  /** Marks the unit's first calculation done (requests after it are edits). */
  markCalculated(unitId: string): void {
    this.calculated.add(unitId)
  }

  /** A user edit of the formula cell: retyping the same formula runs it again. */
  forgetCell(unitId: string, sheetId: string, row: number, column: number): void {
    this.done.delete(ownerKey({ unitId, sheetId, row, column }))
  }

  /** Remembered requests of one function (the add-in's cache entries). */
  count(fn: string): number {
    let n = 0
    for (const entry of this.done.values()) if (entry.fn === fn) n++
    return n
  }

  /** Forgets one function's fingerprints so its formulas re-apply on their next calculation. */
  forget(fn: string): number {
    let n = 0
    for (const [key, entry] of [...this.done]) {
      if (entry.fn !== fn) continue
      this.done.delete(key)
      n++
    }
    return n
  }

  clearUnit(unitId: string): void {
    this.calculated.delete(unitId)
    for (const key of [...this.done.keys()])
      if (key.startsWith(`${unitId}\u0000`)) this.done.delete(key)
    for (const key of [...this.pending.keys()])
      if (key.startsWith(`${unitId}\u0000`)) this.pending.delete(key)
  }
}

export const dvhCommands = new DvhCommandQueue()

// ---------------- pure helpers ----------------

/** Excel's defined-name rules (no cell-like names such as A1 or R1C1). */
export function isValidDefinedName(name: string): boolean {
  if (name.length === 0 || name.length > 255) return false
  if (!/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name)) return false
  if (/^[A-Za-z]{1,3}\d{1,7}$/.test(name)) return false
  if (/^[RrCc]$/.test(name) || /^[Rr]\d*[Cc]\d*$/.test(name)) return false
  return true
}

/** `Sheet1` → `Sheet1`, `Kế hoạch 1` → `'Kế hoạch 1'`. */
export function quoteSheetName(name: string): string {
  return /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`
}

export function absoluteAddress(b: Bounds): string {
  const start = `$${columnLabel(b.startColumn)}$${b.startRow + 1}`
  return b.startRow === b.endRow && b.startColumn === b.endColumn
    ? start
    : `${start}:$${columnLabel(b.endColumn)}$${b.endRow + 1}`
}

export function relativeAddress(b: Bounds): string {
  return absoluteAddress(b).replace(/\$/g, '')
}

const A1_REFERENCE =
  /^((?:'[^']+'|[\p{L}_][\p{L}\p{N}_.]*)!)?\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$/u

/** DelName's matching: exact name (case-insensitive) or a `*`/`?` wildcard. */
export function wildcardMatch(name: string, pattern: string): boolean {
  const source = pattern
    .split('')
    .map((ch) => (ch === '*' ? '.*' : ch === '?' ? '.' : ch.replace(/[.+^${}()|[\]\\]/g, '\\$&')))
    .join('')
  return new RegExp(`^${source}$`, 'i').test(name)
}

export function namesToDelete(
  names: readonly { name: string; ref: string; hidden: boolean }[],
  option: string,
): string[] {
  const upper = option.trim().toUpperCase()
  if (['ALL', 'ERR', 'OUT', 'HIDDEN'].includes(upper))
    return names.filter((info) => nameMatchesCategory(info, upper)).map((info) => info.name)
  return names
    .filter(
      (info) =>
        info.name.toLowerCase() === option.trim().toLowerCase() ||
        ((option.includes('*') || option.includes('?')) && wildcardMatch(info.name, option.trim())),
    )
    .map((info) => info.name)
}

/** "[a, b, c, d, e... +3 more]" as the add-in listed names. */
export function nameListPreview(names: readonly string[]): string {
  const extra = names.length > 5 ? `... +${names.length - 5} more` : ''
  return `[${names.slice(0, 5).join(', ')}${extra}]`
}

/** The cell font as a fully explicit rich run (a rich cell no longer inherits it). */
export function fontRun(text: string, style: Nullable<IStyleData>): WorkbookRichRun {
  return {
    text,
    bold: style?.bl === 1,
    italic: style?.it === 1,
    underline: style?.ul?.s === 1,
    strikethrough: style?.st?.s === 1,
    ...(typeof style?.cl?.rgb === 'string' ? { color: style.cl.rgb } : {}),
    ...(typeof style?.fs === 'number' ? { size: style.fs } : {}),
    ...(typeof style?.ff === 'string' ? { family: style.ff } : {}),
  }
}

/** The cell's runs, each carrying the full font (cell font under the run's own settings). */
export function cellRuns(
  text: string,
  style: Nullable<IStyleData>,
  rich: readonly WorkbookRichRun[] | undefined,
): WorkbookRichRun[] {
  const base = fontRun('', style)
  if (!rich || rich.length === 0) return text === '' ? [] : [{ ...base, text }]
  return rich.map((run) => {
    const styled =
      run.bold ||
      run.italic ||
      run.underline ||
      run.strikethrough ||
      run.color !== undefined ||
      run.size !== undefined ||
      run.family !== undefined ||
      run.vertAlign !== undefined
    return styled ? { ...base, ...run } : { ...base, text: run.text }
  })
}

/** JoinFormat's orders: 1 left→right top→down, 2 bottom→up, 3 right→left, 4 both reversed. */
export function orderJoinCells<T extends { row: number; column: number }>(
  cells: readonly T[],
  style: number,
): T[] {
  const sorted = [...cells]
  const rowSign = style === 2 || style === 4 ? -1 : 1
  const columnSign = style === 3 || style === 4 ? -1 : 1
  if (style < 1 || style > 4) return sorted
  return sorted.sort((a, b) => rowSign * (a.row - b.row) || columnSign * (a.column - b.column))
}

/** Joins texts with their runs; the delimiter carries no formatting of its own. */
export function joinRichParts(
  parts: readonly { text: string; runs: readonly WorkbookRichRun[] }[],
  delimiter: string,
): { text: string; runs: WorkbookRichRun[] } {
  let text = ''
  const runs: WorkbookRichRun[] = []
  for (const part of parts) {
    if (part.text === '') continue
    if (text !== '' && delimiter !== '') {
      text += delimiter
      runs.push({
        text: delimiter,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
      })
    }
    text += part.text
    runs.push(...part.runs)
  }
  return { text, runs }
}

export type RunFormat = Partial<Omit<WorkbookRichRun, 'text'>>

/**
 * FindFormat on one cell: every case-insensitive occurrence of each pattern
 * gets `format` on top of the cell's own runs. Null when nothing matches.
 */
export function highlightMatches(
  runs: readonly WorkbookRichRun[],
  patterns: readonly string[],
  format: RunFormat,
): { runs: WorkbookRichRun[]; matches: number } | null {
  const text = runs.map((run) => run.text).join('')
  const lower = text.toLowerCase()
  const marked = new Array<boolean>(text.length).fill(false)
  let matches = 0
  for (const pattern of patterns) {
    if (pattern === '') continue
    const needle = pattern.toLowerCase()
    let from = 0
    for (;;) {
      const at = lower.indexOf(needle, from)
      if (at < 0) break
      for (let i = at; i < at + needle.length; i++) marked[i] = true
      matches++
      from = at + needle.length
    }
  }
  if (matches === 0) return null
  const out: WorkbookRichRun[] = []
  let offset = 0
  for (const run of runs) {
    let start = 0
    while (start < run.text.length) {
      const flag = marked[offset + start]!
      let end = start + 1
      while (end < run.text.length && marked[offset + end] === flag) end++
      const segment = run.text.slice(start, end)
      out.push(flag ? { ...run, ...format, text: segment } : { ...run, text: segment })
      start = end
    }
    offset += run.text.length
  }
  return { runs: out, matches }
}

// ---------------- functions ----------------

type CellRange = ReturnType<FWorksheet['getRange']>

/** A cell's shown text and its runs; a rich cell's text comes from its document, not the display string. */
function cellTextAndRuns(range: CellRange): { text: string; runs: WorkbookRichRun[] } {
  const data = range.getCellData()
  const rich = data?.p ? extractRichText(data.p) : undefined
  const text = rich ? rich.text : (range.getDisplayValue() ?? '').replace(/\r?\n$/, '')
  return { text, runs: cellRuns(text, range.getCellStyleData(), rich?.runs) }
}

interface CommandContext {
  readonly runtime: UniverRuntime
  readonly queue: DvhCommandQueue
}

const str = (value: BaseValueObject | undefined, fallback = ''): string => {
  if (isBlankArg(value)) return fallback
  const scalar = scalarOf(value)
  return scalar === null ? fallback : String(scalar)
}

const flag = (value: BaseValueObject | undefined, fallback: boolean): boolean =>
  isBlankArg(value) ? fallback : !isFalseLike(scalarOf(value))

abstract class DvhCommandFunction extends BaseFunction {
  override needsReferenceObject = true
  constructor(
    name: string,
    protected readonly ctx: CommandContext,
  ) {
    super(name.toUpperCase())
  }
  protected workbook() {
    const workbook = this.ctx.runtime.univerAPI.getActiveWorkbook()
    return workbook && (!this.unitId || workbook.getId() === this.unitId) ? workbook : null
  }
  protected sheet(sheetId: string): FWorksheet | null {
    return this.workbook()?.getSheetBySheetId(sheetId) ?? null
  }
  protected submit(key: string, run: CommandRequest['run']): void {
    if (!this.unitId || !this.subUnitId) return
    this.ctx.queue.request({
      unitId: this.unitId,
      sheetId: this.subUnitId,
      row: this.row,
      column: this.column,
      fn: String(this.name),
      key,
      run,
    })
  }
  protected text(message: string): BaseValueObject {
    return StringValueObject.create(message)
  }
  protected referenceAddress(
    reference: BaseReferenceObject,
    b: Bounds = boundsOf(reference),
  ): string {
    const name = this.sheet(reference.getSheetId())?.getSheetName() ?? ''
    return `${quoteSheetName(name)}!${absoluteAddress(b)}`
  }
}

class DvhAddName extends DvhCommandFunction {
  override minParams = 1
  override maxParams = 5
  constructor(ctx: CommandContext) {
    super('DVH.AddName', ctx)
  }
  override calculate(
    nameInput: BaseValueObject,
    refersTo?: BaseValueObject,
    scopeArg?: BaseValueObject,
    commentArg?: BaseValueObject,
    turboArg?: BaseValueObject,
  ): BaseValueObject {
    const mode = flag(turboArg, true) ? 'TURBO' : 'NORMAL'
    const scope = Math.trunc(Number(scalarOf(scopeArg) ?? 0)) === 1 ? 1 : 0
    const comment = str(commentArg)
    const nameRef = nameInput.isReferenceObject() ? asReference(nameInput) : null
    const nameBounds = nameRef ? boundsOf(nameRef) : null
    const names = matrixOf(nameInput)
    const nameMatrix: Scalar[][] = names === null ? [] : Array.isArray(names) ? names : [[names]]
    if (nameMatrix.length === 0) return this.text('ERROR: No names provided')
    const refRef = refersTo?.isReferenceObject() ? asReference(refersTo) : null
    const refBounds = refRef ? boundsOf(refRef) : null
    const refValues = !isBlankArg(refersTo) && !refRef ? matrixOf(refersTo) : null
    const refMatrix: Scalar[][] | null =
      refValues === null ? null : Array.isArray(refValues) ? refValues : [[refValues]]
    const pairs: { name: string; ref: string; error?: string }[] = []
    nameMatrix.forEach((row, i) =>
      row.forEach((cell, j) => {
        const name = String(cell ?? '').trim()
        if (name === '') return
        let ref: string | null = null
        let error: string | undefined
        if (refRef && refBounds) {
          const sameShape =
            nameBounds &&
            nameBounds.endRow - nameBounds.startRow === refBounds.endRow - refBounds.startRow &&
            nameBounds.endColumn - nameBounds.startColumn ===
              refBounds.endColumn - refBounds.startColumn
          if (sameShape) {
            const row0 = Math.min(refBounds.startRow + i, refBounds.endRow)
            const col0 = Math.min(refBounds.startColumn + j, refBounds.endColumn)
            ref = this.referenceAddress(refRef, {
              startRow: row0,
              endRow: row0,
              startColumn: col0,
              endColumn: col0,
            })
          } else ref = this.referenceAddress(refRef, refBounds)
        } else if (refMatrix) {
          const value = String(refMatrix[i]?.[j] ?? refMatrix[0]?.[0] ?? '').trim()
          if (value.startsWith('=') || A1_REFERENCE.test(value)) ref = value
          else if (value !== '') error = `'${value}' is not a valid Excel address/formula`
        } else if (nameRef && nameBounds) {
          const r = nameBounds.startRow + i
          const c = nameBounds.startColumn + j
          ref = this.referenceAddress(nameRef, {
            startRow: r,
            endRow: r,
            startColumn: c,
            endColumn: c,
          })
        } else {
          const sheetName = this.sheet(this.subUnitId ?? '')?.getSheetName() ?? ''
          ref = `${quoteSheetName(sheetName)}!${absoluteAddress({ startRow: this.row, endRow: this.row, startColumn: this.column, endColumn: this.column })}`
        }
        if (!error && !isValidDefinedName(name)) error = 'invalid name'
        if (!error && !ref) error = 'Reference address is empty'
        pairs.push({ name, ref: (ref ?? '').replace(/^=/, ''), ...(error ? { error } : {}) })
      }),
    )
    if (pairs.length === 0) return this.text('ERROR: No valid names to create')
    const valid = pairs.filter((p) => !p.error)
    const errors = pairs.filter((p) => p.error).map((p) => `'${p.name}': ${p.error}`)
    if (valid.length > 0) {
      const sheetId = this.subUnitId ?? ''
      this.submit(JSON.stringify({ valid, scope, comment, sheetId: scope ? sheetId : '' }), () => {
        const workbook = this.workbook()
        if (!workbook) return
        const scopeSheet = scope ? workbook.getSheetBySheetId(sheetId) : null
        for (const pair of valid) {
          const existing = workbook
            .getDefinedNames()
            .find(
              (d) =>
                d.getName().toLowerCase() === pair.name.toLowerCase() &&
                (scopeSheet ? d.getLocalSheetId() === sheetId : d.isWorkbookScope()),
            )
          let builder = existing
            ? existing.toBuilder()
            : workbook.newDefinedNameBuilder().setName(pair.name)
          builder = builder.setFormula(pair.ref)
          if (comment) builder = builder.setComment(comment)
          builder = scopeSheet
            ? builder.setScopeToWorksheet(scopeSheet)
            : builder.setScopeToWorkbook()
          if (existing) workbook.updateDefinedNameBuilder(builder.build())
          else workbook.insertDefinedNameBuilder(builder.build())
        }
        return `DVH.AddName: ${valid.length} Name`
      })
    }
    if (errors.length === 0) return this.text(`✅ ${mode} AddName: ${valid.length} created`)
    if (valid.length > 0)
      return this.text(
        `⚠️ ${mode} PARTIAL: ${valid.length} OK, ${errors.length} Failed\nErrors: ${errors.slice(0, 2).join('; ')}`,
      )
    return this.text(`❌ ${mode} FAILED: ${errors.slice(0, 3).join('; ')}`)
  }
}

class DvhDelName extends DvhCommandFunction {
  override minParams = 1
  override maxParams = 2
  constructor(ctx: CommandContext) {
    super('DVH.DelName', ctx)
  }
  override calculate(optionArg: BaseValueObject, turboArg?: BaseValueObject): BaseValueObject {
    const mode = flag(turboArg, true) ? 'TURBO' : 'NORMAL'
    const option = str(optionArg).trim()
    const names = workbookNames(this.ctx.runtime, this.unitId ?? '')
    if (names.length === 0) return this.text('No names to process')
    const matched = namesToDelete(names, option)
    if (matched.length === 0) return this.text(`No names found matching '${option}'`)
    this.submit(option.toUpperCase(), () => {
      const workbook = this.workbook()
      if (!workbook) return
      const current = namesToDelete(workbookNames(this.ctx.runtime, workbook.getId()), option)
      const doomed = new Set(current.map((name) => name.toLowerCase()))
      let deleted = 0
      for (const defined of workbook.getDefinedNames()) {
        const name = defined.getName()
        if (isDvhDefinedName(name) || !doomed.has(name.toLowerCase())) continue
        defined.delete()
        deleted++
      }
      return `DVH.DelName: ${deleted} Name`
    })
    return this.text(
      `✅ ${mode} DelName: ${matched.length}/${matched.length} deleted\n${nameListPreview(matched)}`,
    )
  }
}

class DvhShowName extends DvhCommandFunction {
  override minParams = 0
  override maxParams = 1
  constructor(ctx: CommandContext) {
    super('DVH.ShowName', ctx)
  }
  override calculate(turboArg?: BaseValueObject): BaseValueObject {
    const mode = flag(turboArg, true) ? 'TURBO' : 'NORMAL'
    const names = workbookNames(this.ctx.runtime, this.unitId ?? '')
    if (names.length === 0) return this.text('No names in workbook')
    const hidden = names.filter((info) => info.hidden).map((info) => info.name)
    if (hidden.length === 0) return this.text('No hidden names found')
    this.submit(JSON.stringify(hidden), () => {
      const workbook = this.workbook()
      if (!workbook) return
      const wanted = new Set(hidden.map((name) => name.toLowerCase()))
      for (const defined of workbook.getDefinedNames()) {
        if (wanted.has(defined.getName().toLowerCase())) defined.setHidden(false)
      }
      return `DVH.ShowName: ${hidden.length} Name`
    })
    return this.text(`✅ ${mode} ShowName: ${hidden.length} unhidden\n${nameListPreview(hidden)}`)
  }
}

type NoteSheet = FWorksheet & { getNotes(): { row: number; col: number; note: string }[] }
type NoteRange = {
  createOrUpdateNote(note: {
    id: string
    row: number
    col: number
    width: number
    height: number
    note: string
  }): unknown
  deleteNote(): unknown
}

class DvhTextCom extends DvhCommandFunction {
  override minParams = 2
  override maxParams = 3
  constructor(ctx: CommandContext) {
    super('DVH.TextCom', ctx)
  }
  override calculate(
    target: BaseValueObject,
    textArg: BaseValueObject,
    titleArg?: BaseValueObject,
  ): BaseValueObject {
    if (!target.isReferenceObject()) return this.text('#ERROR: Cannot determine target range')
    const text = str(textArg)
    if (text.trim() === '')
      return this.text('#ERROR: Comment text is required / Nội dung chú thích bắt buộc')
    if (text.length > 32_000)
      return this.text('⚠️Text truncated to 32,000 characters / Văn bản bị cắt xuống 32.000 ký tự')
    const title = isBlankArg(titleArg) ? 'DVH-Addin' : str(titleArg)
    const note = title.trim() === '' ? text : `${title}\n${text}`
    const reference = asReference(target)
    const b = boundsOf(reference)
    const sheetId = reference.getSheetId()
    const address = `${columnLabel(b.startColumn)}${b.startRow + 1}`
    this.submit(JSON.stringify([sheetId, b.startRow, b.startColumn, note]), () => {
      const sheet = this.sheet(sheetId)
      if (!sheet) return
      ;(sheet.getRange(b.startRow, b.startColumn, 1, 1) as unknown as NoteRange).createOrUpdateNote(
        {
          id: `note-${sheetId}-${b.startRow}-${b.startColumn}`,
          row: b.startRow,
          col: b.startColumn,
          width: 220,
          height: 90,
          note,
        },
      )
      return `DVH.TextCom → ${address}`
    })
    return this.text(`✓ Comment added to ${address}`)
  }
}

class DvhDelCom extends DvhCommandFunction {
  override minParams = 1
  override maxParams = 1
  constructor(ctx: CommandContext) {
    super('DVH.DelCom', ctx)
  }
  override calculate(target: BaseValueObject): BaseValueObject {
    if (!target.isReferenceObject()) return this.text('#ERROR: Cannot determine target range')
    const reference = asReference(target)
    const b = boundsOf(reference)
    const sheetId = reference.getSheetId()
    const cells = (b.endRow - b.startRow + 1) * (b.endColumn - b.startColumn + 1)
    this.submit(JSON.stringify([sheetId, b]), () => {
      const sheet = this.sheet(sheetId) as NoteSheet | null
      if (!sheet) return
      const notes = sheet
        .getNotes()
        .filter(
          (n) =>
            n.row >= b.startRow &&
            n.row <= b.endRow &&
            n.col >= b.startColumn &&
            n.col <= b.endColumn,
        )
      for (const n of notes)
        (sheet.getRange(n.row, n.col, 1, 1) as unknown as NoteRange).deleteNote()
      return `DVH.DelCom: ${notes.length}`
    })
    return this.text(
      `✓ Comments cleared from ${relativeAddress(b)} (${cells.toLocaleString('en-US')} cells)`,
    )
  }
}

class DvhPrintArea extends DvhCommandFunction {
  override minParams = 0
  override maxParams = 255
  constructor(
    ctx: CommandContext,
    private readonly host: () => DvhCommandHost | null,
  ) {
    super('DVH.PrintArea', ctx)
  }
  override calculate(...areas: BaseValueObject[]): BaseValueObject {
    const addresses: string[] = []
    for (const area of areas) {
      if (area.isReferenceObject()) addresses.push(relativeAddress(boundsOf(asReference(area))))
      else {
        const text = str(area).trim()
        if (text !== '') addresses.push(text.replace(/^.*!/, '').replace(/\$/g, ''))
      }
    }
    const sheetId = this.subUnitId ?? ''
    if (addresses.some((a) => !/^[A-Za-z]{1,3}\d{1,7}(:[A-Za-z]{1,3}\d{1,7})?$/.test(a)))
      return this.text('#ERROR: Invalid area')
    const first = addresses[0]?.toUpperCase() ?? null
    this.submit(JSON.stringify(first), () => {
      if (!this.host()?.recordPrintSetup(sheetId, { printArea: first }))
        return 'DVH.PrintArea: cần mở tệp XLSX'
      return first ? `DVH.PrintArea: ${first}` : 'DVH.PrintArea: đã xóa'
    })
    if (!first) return this.text('Cleared\nĐã xóa vùng in')
    const shown = addresses
      .map((a) => a.toUpperCase().replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'))
      .join(',')
    const note = addresses.length > 1 ? '\n(DVH Office in vùng đầu tiên)' : ''
    return this.text(`Set: ${shown}\nĐã thiết lập: ${shown}${note}`)
  }
}

class DvhPrintTitle extends DvhCommandFunction {
  override minParams = 0
  override maxParams = 2
  constructor(
    ctx: CommandContext,
    private readonly host: () => DvhCommandHost | null,
  ) {
    super('DVH.PrintTitle', ctx)
  }
  override calculate(range?: BaseValueObject, isRowArg?: BaseValueObject): BaseValueObject {
    const sheetId = this.subUnitId ?? ''
    if (isBlankArg(range)) {
      this.submit('clear', () => {
        if (!this.host()?.recordPrintSetup(sheetId, { printTitles: null }))
          return 'DVH.PrintTitle: cần mở tệp XLSX'
        return 'DVH.PrintTitle: đã xóa'
      })
      return this.text('Cleared Titles\nĐã xóa tiêu đề in (Dòng & Cột)')
    }
    if (!range.isReferenceObject()) return this.text('Invalid Range\nVùng không hợp lệ')
    const b = boundsOf(asReference(range))
    if (!flag(isRowArg, true)) {
      const cols = `$${columnLabel(b.startColumn)}:$${columnLabel(b.endColumn)}`
      return this.text(`Cols: ${cols}\nDVH Office chưa hỗ trợ cột tiêu đề khi in`)
    }
    const rows = `${b.startRow + 1}:${b.endRow + 1}`
    this.submit(rows, () => {
      if (!this.host()?.recordPrintSetup(sheetId, { printTitles: rows }))
        return 'DVH.PrintTitle: cần mở tệp XLSX'
      return `DVH.PrintTitle: ${rows}`
    })
    const shown = `$${b.startRow + 1}:$${b.endRow + 1}`
    return this.text(`Rows: ${shown}\nĐã đặt dòng: ${shown}`)
  }
}

class DvhFitRow extends DvhCommandFunction {
  override minParams = 1
  override maxParams = 2
  constructor(ctx: CommandContext) {
    super('DVH.FitRow', ctx)
  }
  override calculate(range: BaseValueObject, clearArg?: BaseValueObject): BaseValueObject {
    if (!range.isReferenceObject()) return this.text('❌ Range phải là vùng chọn hợp lệ')
    const reference = asReference(range)
    const b = boundsOf(reference)
    const rows = b.endRow - b.startRow + 1
    const columns = b.endColumn - b.startColumn + 1
    if (rows * columns > MAX_COMMAND_CELLS) return this.text('[❌ Error: Range too large]')
    const clear = flag(clearArg, false)
    const sheetId = reference.getSheetId()
    const message = clear ? `✅ ${rows} rows [Formulas Cleared]` : `✅ ${rows} rows`
    const owner = { sheetId: this.subUnitId ?? '', row: this.row, column: this.column }
    // fingerprint the text: a longer value must refit the rows
    const texts = matrixOf(range)
    this.submit(JSON.stringify([sheetId, b, clear, texts]), () => {
      const sheet = this.sheet(sheetId)
      if (!sheet) return
      const target = sheet.getRange(b.startRow, b.startColumn, rows, columns)
      fitRowHeights(sheet, b.startRow, b.startColumn, rows, columns, target.getCellDatas())
      if (clear) {
        // formulas become values, then the formula removes itself (the add-in's self-destruct)
        target.setValues(
          target.getValues().map((row) => row.map((v) => ({ v: v ?? null, f: null, si: null }))),
        )
        this.sheet(owner.sheetId)
          ?.getRange(owner.row, owner.column, 1, 1)
          .setValues([[{ v: message, f: null, si: null }]])
      }
      return `DVH.FitRow: ${rows}`
    })
    return this.text(message)
  }
}

class DvhJoinFormat extends DvhCommandFunction {
  override minParams = 2
  override maxParams = 255
  constructor(ctx: CommandContext) {
    super('DVH.JoinFormat', ctx)
  }
  override calculate(
    target: BaseValueObject,
    delimiterArg?: BaseValueObject,
    styleArg?: BaseValueObject,
    ...ranges: BaseValueObject[]
  ): BaseValueObject {
    if (!target.isReferenceObject()) return ErrorRef()
    const delimiter = str(delimiterArg)
    const style = Math.trunc(Number(scalarOf(styleArg) ?? 1)) || 1
    const sources = ranges.filter((range) => range.isReferenceObject()).map(asReference)
    if (sources.length === 0) return this.text('#ERROR: No source ranges')
    const targetRef = asReference(target)
    const tb = boundsOf(targetRef)
    const targetSheetId = targetRef.getSheetId()
    if (
      targetSheetId === this.subUnitId &&
      tb.startRow === this.row &&
      tb.startColumn === this.column
    )
      return this.text('#ERROR: target_cell không được là ô chứa công thức')
    const cells: { sheet: FWorksheet; row: number; column: number }[] = []
    for (const source of sources) {
      const sheet = this.sheet(source.getSheetId())
      if (!sheet) continue
      const b = boundsOf(source)
      if (
        (b.endRow - b.startRow + 1) * (b.endColumn - b.startColumn + 1) + cells.length >
        MAX_COMMAND_CELLS
      )
        return this.text('#ERROR: Range too large')
      for (let row = b.startRow; row <= b.endRow; row++)
        for (let column = b.startColumn; column <= b.endColumn; column++)
          cells.push({ sheet, row, column })
    }
    const parts = orderJoinCells(cells, style).map(({ sheet, row, column }) => {
      return cellTextAndRuns(sheet.getRange(row, column, 1, 1))
    })
    const joined = joinRichParts(parts, delimiter)
    this.submit(JSON.stringify([targetSheetId, tb.startRow, tb.startColumn, joined]), () => {
      const sheet = this.sheet(targetSheetId)
      if (!sheet) return
      sheet
        .getRange(tb.startRow, tb.startColumn, 1, 1)
        .setValues([
          [
            joined.text === ''
              ? { v: null, p: null, f: null, si: null }
              : ({
                  v: joined.text,
                  p: toRichTextDocument(joined.text, joined.runs),
                  f: null,
                  si: null,
                } as ICellData),
          ],
        ])
      return `DVH.JoinFormat → ${columnLabel(tb.startColumn)}${tb.startRow + 1}`
    })
    return this.text(joined.text)
  }
}

const ErrorRef = (): BaseValueObject => StringValueObject.create('#REF!')

class DvhFindFormat extends DvhCommandFunction {
  override minParams = 2
  override maxParams = 3
  constructor(ctx: CommandContext) {
    super('DVH.FindFormat', ctx)
  }
  override calculate(
    searchArg: BaseValueObject,
    findArg: BaseValueObject,
    formatArg?: BaseValueObject,
  ): BaseValueObject {
    if (!searchArg.isReferenceObject()) return ErrorRef()
    const search = asReference(searchArg)
    const b = boundsOf(search)
    if ((b.endRow - b.startRow + 1) * (b.endColumn - b.startColumn + 1) > MAX_COMMAND_CELLS)
      return this.text('#ERROR: Range too large')
    const found = matrixOf(findArg)
    const patterns = (Array.isArray(found) ? found.flat() : found === null ? [] : [found])
      .map((value) => String(value ?? ''))
      .filter((text) => text !== '')
    if (patterns.length === 0) return this.text('#NO_PATTERNS')
    let format: RunFormat = { bold: true, color: '#FF0000' }
    if (formatArg?.isReferenceObject()) {
      const ref = asReference(formatArg)
      const fb = boundsOf(ref)
      const cell = this.sheet(ref.getSheetId())?.getRange(fb.startRow, fb.startColumn, 1, 1)
      if (cell && (cell.getDisplayValue() ?? '') !== '') {
        const { text: _text, ...spec } = fontRun('', cell.getCellStyleData())
        format = spec
      }
    }
    const sheetId = search.getSheetId()
    const sheet = this.sheet(sheetId)
    if (!sheet) return ErrorRef()
    const edits: { row: number; column: number; text: string; runs: WorkbookRichRun[] }[] = []
    let total = 0
    for (let row = b.startRow; row <= b.endRow; row++) {
      for (let column = b.startColumn; column <= b.endColumn; column++) {
        const range = sheet.getRange(row, column, 1, 1)
        const data = range.getCellData()
        // character formats cannot live on a formula cell, as in Excel
        if (!data || data.f || data.si || data.v == null || data.v === '') continue
        const { text, runs } = cellTextAndRuns(range)
        if (text === '') continue
        const result = highlightMatches(runs, patterns, format)
        if (!result) continue
        total += result.matches
        edits.push({ row, column, text, runs: result.runs })
      }
    }
    this.submit(
      JSON.stringify([sheetId, b, patterns, format, edits.map((e) => [e.row, e.column, e.text])]),
      () => {
        const target = this.sheet(sheetId)
        if (!target) return
        for (const edit of edits) {
          target
            .getRange(edit.row, edit.column, 1, 1)
            .setValues([[{ p: toRichTextDocument(edit.text, edit.runs) } as ICellData]])
        }
        return `DVH.FindFormat: ${total} matches`
      },
    )
    return this.text(`✓ ${total} matches`)
  }
}

/** DVH.Barcode types 0-8 as bwip-js encoders. */
export const BARCODE_ENCODERS = [
  'qrcode',
  'datamatrix',
  'azteccode',
  'maxicode',
  'code128',
  'ean13',
  'upca',
  'interleaved2of5',
  'code39',
] as const

/** Renders a barcode to a PNG data URL (bwip-js loads on first use). */
async function renderBarcode(
  content: string,
  type: number,
): Promise<{ dataUrl: string; mediaType: string; width: number; height: number }> {
  const bwip = await import('bwip-js/browser')
  const canvas = document.createElement('canvas')
  const encoder = BARCODE_ENCODERS[type]!
  const twoDimensional = type <= 3
  bwip.toCanvas(canvas, {
    bcid: encoder,
    text: content,
    scale: twoDimensional ? 3 : 2,
    ...(twoDimensional ? {} : { height: 12, includetext: true, textxalign: 'center' }),
    backgroundcolor: 'FFFFFF',
    paddingwidth: 2,
    paddingheight: 2,
  })
  return {
    dataUrl: canvas.toDataURL('image/png'),
    mediaType: 'image/png',
    width: canvas.width,
    height: canvas.height,
  }
}

class DvhBarcode extends DvhCommandFunction {
  override minParams = 1
  override maxParams = 4
  constructor(
    ctx: CommandContext,
    private readonly host: () => DvhCommandHost | null,
  ) {
    super('DVH.Barcode', ctx)
  }
  override calculate(
    contentArg: BaseValueObject,
    targetArg?: BaseValueObject,
    typeArg?: BaseValueObject,
    showTextArg?: BaseValueObject,
  ): BaseValueObject {
    const content = str(contentArg)
    if (content.trim() === '')
      return this.text('⚠️ ERROR: Content cannot be empty / Nội dung không được rỗng')
    const type = isBlankArg(typeArg) ? 0 : Number(scalarOf(typeArg))
    if (!Number.isInteger(type) || type < 0 || type > 8)
      return this.text(
        '⚠️ ERROR: Invalid barcode type. Must be 0-8 / Loại mã không hợp lệ. Phải từ 0-8',
      )
    let sheetId = this.subUnitId ?? ''
    let row = this.row
    let column = this.column
    if (targetArg?.isReferenceObject()) {
      const reference = asReference(targetArg)
      const b = boundsOf(reference)
      sheetId = reference.getSheetId()
      row = b.startRow
      column = b.startColumn
    }
    const name = `DVH_Barcode_${sheetId}_${row}_${column}`
    this.submit(JSON.stringify([content, type, sheetId, row, column]), async () => {
      const image = await renderBarcode(content, type)
      if (!this.host()?.placePicture(sheetId, row, column, name, image)) return
      return `DVH.Barcode → ${columnLabel(column)}${row + 1}`
    })
    return this.text(flag(showTextArg, false) ? content : '✓')
  }
}

/** ClearFindCache, FindCacheStats, FindFormatStatus, ClearFormatCache. */
class DvhCacheFunction extends BaseFunction {
  override minParams = 0
  override maxParams = 0
  constructor(
    name: string,
    private readonly queue: DvhCommandQueue,
    private readonly answer: (queue: DvhCommandQueue) => string,
  ) {
    super(name.toUpperCase())
  }
  override calculate(): BaseValueObject {
    return StringValueObject.create(this.answer(this.queue))
  }
}

const FIND = 'DVH.FINDFORMAT'
const JOIN = 'DVH.JOINFORMAT'
const cacheStats = (q: DvhCommandQueue): string => `Cache: ${q.count(FIND)} entries, 0 format specs`

const descriptions: Pick<IFunctionInfo, 'functionName' | 'abstract' | 'functionParameter'>[] = [
  {
    functionName: 'DVH.ADDNAME',
    abstract: 'Tạo Name (một hoặc hàng loạt) khi nhập công thức',
    functionParameter: [
      {
        name: 'name_input',
        detail: 'Tên Name hoặc vùng chứa danh sách tên',
        example: '"DonGia"',
        require: 1,
        repeat: 0,
      },
      {
        name: 'refers_to',
        detail: 'Vùng/địa chỉ tham chiếu (để trống: chính ô tên)',
        example: 'B2:B20',
        require: 0,
        repeat: 0,
      },
      {
        name: 'scope',
        detail: '0 = Workbook (mặc định), 1 = sheet hiện tại',
        example: '0',
        require: 0,
        repeat: 0,
      },
      { name: 'comment', detail: 'Ghi chú cho Name', example: '""', require: 0, repeat: 0 },
      {
        name: 'turbo_mode',
        detail: 'Giữ để tương thích DVH-Tool',
        example: 'TRUE',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.DELNAME',
    abstract: 'Xóa Name: ALL, ERR (lỗi), OUT (liên kết ngoài), HIDDEN, tên hoặc mẫu *',
    functionParameter: [
      {
        name: 'option',
        detail: 'ALL, ERR, OUT, HIDDEN, tên hoặc mẫu (vd "tmp*")',
        example: '"ERR"',
        require: 1,
        repeat: 0,
      },
      {
        name: 'turbo_mode',
        detail: 'Giữ để tương thích DVH-Tool',
        example: 'TRUE',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.SHOWNAME',
    abstract: 'Hiện tất cả Name đang ẩn',
    functionParameter: [
      {
        name: 'turbo_mode',
        detail: 'Giữ để tương thích DVH-Tool',
        example: 'TRUE',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.TEXTCOM',
    abstract: 'Thêm chú thích (note) vào ô',
    functionParameter: [
      { name: 'target_cell', detail: 'Ô nhận chú thích', example: 'B5', require: 1, repeat: 0 },
      {
        name: 'comment_text',
        detail: 'Nội dung chú thích',
        example: '"Đã kiểm tra"',
        require: 1,
        repeat: 0,
      },
      {
        name: 'title',
        detail: 'Dòng tiêu đề (mặc định DVH-Addin)',
        example: '"KS"',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.DELCOM',
    abstract: 'Xóa mọi chú thích (note) trong vùng',
    functionParameter: [
      {
        name: 'target_range',
        detail: 'Vùng cần xóa chú thích',
        example: 'A1:H50',
        require: 1,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.PRINTAREA',
    abstract: 'Đặt vùng in cho sheet chứa công thức (để trống: xóa vùng in)',
    functionParameter: [
      { name: 'areas', detail: 'Vùng in', example: 'A1:H60', require: 0, repeat: 1 },
    ],
  },
  {
    functionName: 'DVH.PRINTTITLE',
    abstract: 'Đặt dòng tiêu đề lặp lại khi in (để trống: xóa)',
    functionParameter: [
      { name: 'title_range', detail: 'Vùng tiêu đề', example: 'A1:H3', require: 0, repeat: 0 },
      {
        name: 'is_row',
        detail: 'TRUE = dòng tiêu đề (mặc định)',
        example: 'TRUE',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.FITROW',
    abstract: 'Tự điều chỉnh chiều cao dòng theo nội dung (kể cả ô gộp)',
    functionParameter: [
      { name: 'range', detail: 'Vùng cần điều chỉnh', example: 'A5:H100', require: 1, repeat: 0 },
      {
        name: 'clear_formula',
        detail: 'TRUE = đổi công thức thành giá trị rồi tự xóa',
        example: 'FALSE',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.JOINFORMAT',
    abstract: 'Nối nội dung các ô vào ô đích, giữ định dạng chữ của từng ô',
    functionParameter: [
      { name: 'target_cell', detail: 'Ô đích', example: 'H2', require: 1, repeat: 0 },
      { name: 'delimiter', detail: 'Ký tự ngăn cách', example: '", "', require: 0, repeat: 0 },
      {
        name: 'style',
        detail: 'Thứ tự nối 1-4 (1: trái→phải, trên→dưới)',
        example: '1',
        require: 0,
        repeat: 0,
      },
      { name: 'ranges', detail: 'Các vùng nguồn', example: 'A2:C2', require: 1, repeat: 1 },
    ],
  },
  {
    functionName: 'DVH.FINDFORMAT',
    abstract: 'Tìm chuỗi trong vùng và tô định dạng (mặc định đỏ đậm) cho phần khớp',
    functionParameter: [
      { name: 'search_range', detail: 'Vùng tìm kiếm', example: 'A1:A200', require: 1, repeat: 0 },
      {
        name: 'find_values',
        detail: 'Chuỗi hoặc vùng chứa các chuỗi cần tìm',
        example: '"bê tông"',
        require: 1,
        repeat: 0,
      },
      { name: 'format_cell', detail: 'Ô mẫu định dạng chữ', example: 'K1', require: 0, repeat: 0 },
    ],
  },
  {
    functionName: 'DVH.BARCODE',
    abstract: 'Tạo mã vạch/QR và chèn hình tại ô đích',
    functionParameter: [
      {
        name: 'content',
        detail: 'Nội dung cần mã hóa',
        example: '"DVH-001"',
        require: 1,
        repeat: 0,
      },
      {
        name: 'target_cell',
        detail: 'Ô đặt hình (mặc định: ô công thức)',
        example: 'D2',
        require: 0,
        repeat: 0,
      },
      {
        name: 'barcode_type',
        detail:
          '0 QR, 1 DataMatrix, 2 Aztec, 3 MaxiCode, 4 Code128, 5 EAN13, 6 UPCA, 7 ITF, 8 Code39',
        example: '0',
        require: 0,
        repeat: 0,
      },
      {
        name: 'show_text',
        detail: 'TRUE = ô công thức hiện nội dung',
        example: 'FALSE',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.CLEARFINDCACHE',
    abstract: 'Xóa bộ nhớ của FindFormat (áp lại định dạng ở lần tính sau)',
    functionParameter: [],
  },
  {
    functionName: 'DVH.FINDCACHESTATS',
    abstract: 'Số công thức FindFormat đang được ghi nhớ',
    functionParameter: [],
  },
  {
    functionName: 'DVH.FINDFORMATSTATUS',
    abstract: 'Trạng thái FindFormat',
    functionParameter: [],
  },
  {
    functionName: 'DVH.CLEARFORMATCACHE',
    abstract: 'Xóa bộ nhớ của JoinFormat (nối lại ở lần tính sau)',
    functionParameter: [],
  },
]

export const DVH_COMMAND_FUNCTION_NAMES = descriptions.map((d) => d.functionName)

/**
 * Registers the command functions and runs their queue after each formula
 * calculation. `host` is read lazily: the app's journal changes per file.
 */
export function installDvhFormulaCommands(
  runtime: UniverRuntime,
  host: () => DvhCommandHost | null,
  queue: DvhCommandQueue = dvhCommands,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const functions = injector.get(IFunctionService)
  const ctx: CommandContext = { runtime, queue }
  const executors: BaseFunction[] = [
    new DvhAddName(ctx),
    new DvhDelName(ctx),
    new DvhShowName(ctx),
    new DvhTextCom(ctx),
    new DvhDelCom(ctx),
    new DvhPrintArea(ctx, host),
    new DvhPrintTitle(ctx, host),
    new DvhFitRow(ctx),
    new DvhJoinFormat(ctx),
    new DvhFindFormat(ctx),
    new DvhBarcode(ctx, host),
    new DvhCacheFunction(
      'DVH.ClearFindCache',
      queue,
      (q) => `✓ Cleared ${q.forget(FIND)} entries + 0 format specs`,
    ),
    new DvhCacheFunction('DVH.FindCacheStats', queue, cacheStats),
    new DvhCacheFunction(
      'DVH.FindFormatStatus',
      queue,
      (q) => `✓ Active (DVH Office) | ${cacheStats(q)}`,
    ),
    new DvhCacheFunction(
      'DVH.ClearFormatCache',
      queue,
      (q) => `✓ Cleared ${q.forget(JOIN)} entries`,
    ),
  ]
  functions.registerExecutors(...executors)
  const infos = descriptions.map((d): IFunctionInfo => ({
    ...d,
    functionType: FunctionType.User,
    description: `DVH Tool · ${d.abstract}`,
  }))
  const descriptionHandle = injector.get(IDescriptionService).registerDescriptions(infos)
  const aliases = registerDvhAliases(injector, executors, infos)

  // A user edit of a formula cell lets the same formula run again.
  const commands = injector.get(ICommandService)
  const editListener = commands.onCommandExecuted((command, options) => {
    if (command.id !== SetRangeValuesMutation.id) return
    if ((options as { fromFormula?: boolean } | undefined)?.fromFormula) return
    const p = command.params as {
      unitId: string
      subUnitId: string
      cellValue?: IObjectMatrixPrimitiveType<Nullable<ICellData>>
    }
    if (!p.cellValue) return
    new ObjectMatrix(p.cellValue).forValue((row, column, cell) => {
      if (cell && ('f' in cell || 'si' in cell || 'v' in cell))
        queue.forgetCell(p.unitId, p.subUnitId, row, column)
    })
  })

  let timer: ReturnType<typeof setTimeout> | null = null
  const flush = (): void => {
    timer = null
    const h = host()
    const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
    if (!h) return
    const runs = queue.take((sheetId, row, column) => h.isUserFormula(sheetId, row, column))
    if (unitId && h.isLoaded()) queue.markCalculated(unitId)
    if (runs.length === 0) return
    const messages: string[] = []
    const batch = beginUndoBatch(runtime)
    try {
      for (const request of runs) {
        try {
          const message = request.run()
          if (message instanceof Promise) {
            void message.then(
              (text) => text && h.notify(`✓ ${text}`, 'info'),
              (error: unknown) =>
                h.notify(
                  `${request.fn}: ${error instanceof Error ? error.message : String(error)}`,
                  'error',
                ),
            )
          } else if (message) messages.push(message)
        } catch (error) {
          h.notify(
            `${request.fn}: ${error instanceof Error ? error.message : String(error)}`,
            'error',
          )
        }
      }
    } finally {
      batch.settle()
    }
    if (messages.length > 0) h.notify(`✓ ${messages.join(' · ')} (Ctrl+Z để hoàn tác)`, 'info')
  }
  const calculation = runtime.univerAPI.getFormula().calculationEnd(() => {
    // after the engine has written its results, outside its callback
    timer ??= setTimeout(flush, 0)
  })
  const unitDisposed = injector
    .get(IUniverInstanceService)
    .unitDisposed$.subscribe((unit) => queue.clearUnit(unit.getUnitId()))
  return {
    dispose() {
      if (timer) clearTimeout(timer)
      calculation.dispose()
      editListener.dispose()
      aliases.dispose()
      descriptionHandle.dispose()
      functions.unregisterExecutors(...executors.map((executor) => executor.name))
      unitDisposed.unsubscribe()
    },
  }
}

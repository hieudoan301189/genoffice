/**
 * Format results of DVH worksheet functions (spike S5, see
 * docs/dvh-architecture-implementation-plan.md appendix A).
 *
 * A function publishes the formats it "returns" here, keyed by the formula
 * cell that owns them. An interceptor paints them at render time, the way
 * conditional formatting does: recalculation never mutates the workbook, never
 * adds undo items and cannot loop. Printing reads composed styles through the
 * same interceptor, and save bakes the current formats into real cell styles
 * so Excel shows them too.
 */
import {
  CommandType,
  ICommandService,
  InterceptorEffectEnum,
  IUniverInstanceService,
  ObjectMatrix,
} from '@univerjs/core'
import type {
  ICellData,
  IObjectMatrixPrimitiveType,
  IRange,
  IStyleData,
  IUnitRange,
  Nullable,
  Workbook,
} from '@univerjs/core'
import { IActiveDirtyManagerService } from '@univerjs/engine-formula'
import { IRenderManagerService } from '@univerjs/engine-render'
import {
  ClearSelectionFormatCommand,
  INTERCEPTOR_POINT,
  InsertColMutation,
  InsertRowMutation,
  MoveColsMutation,
  MoveRangeMutation,
  MoveRowsMutation,
  RemoveColMutation,
  RemoveRowMutation,
  RemoveSheetMutation,
  SetBorderCommand,
  SetRangeValuesMutation,
  SetStyleCommand,
  SheetInterceptorService,
} from '@univerjs/sheets'
import { SheetSkeletonManagerService } from '@univerjs/sheets-ui'
import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'

import type { WorkbookCellEdit } from '../shared/desktop-api'
import { fromNeutralStyle } from './edit-journal'
import type { UniverRuntime } from './univer-state'

export interface FormatOwner {
  readonly unitId: string
  readonly sheetId: string
  readonly row: number
  readonly column: number
  /** upper-case function name; the owner stays alive while its cell still calls it */
  readonly functionName: string
}

export interface FormatRect {
  readonly sheetId: string
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
  readonly style: WorkbookStyleEdit
}

export type SourceRect = Omit<FormatRect, 'style'>

export interface BakedStyle {
  readonly sheetId: string
  readonly row: number
  readonly column: number
  readonly style: WorkbookStyleEdit
}

interface OwnerEntry {
  readonly owner: FormatOwner
  readonly rects: readonly FormatRect[]
  readonly signature: string
  /** later publishes win where formats overlap */
  readonly seq: number
}

/** Whole-sheet references (A:A) stay one area; only this many cells get baked per area. */
const MAX_BAKED_CELLS_PER_AREA = 50_000

const ownerKey = (o: Pick<FormatOwner, 'unitId' | 'sheetId' | 'row' | 'column'>): string =>
  `${o.unitId}\u0000${o.sheetId}\u0000${o.row}\u0000${o.column}`
const sheetKey = (unitId: string, sheetId: string): string => `${unitId}\u0000${sheetId}`
const cellKey = (row: number, column: number): string => `${row}:${column}`

export class FormulaFormatStore {
  private readonly owners = new Map<string, OwnerEntry>()
  /** single-cell formats (DVH.Table output) */
  private readonly points = new Map<string, Map<string, Set<string>>>()
  /** multi-cell formats (DVH.Font over a range) */
  private readonly areas = new Map<string, Set<string>>()
  private readonly listeners = new Set<(unitId: string) => void>()
  /** ranges whose formats an owner reads (DVH.Table copies source formats) */
  private readonly sources = new Map<string, { unitId: string; rects: readonly SourceRect[] }>()
  private seq = 0

  /** Replaces everything `owner` publishes; false when nothing changed (no repaint needed). */
  replace(owner: FormatOwner, rects: readonly FormatRect[]): boolean {
    const key = ownerKey(owner)
    const signature = `${owner.functionName}|${JSON.stringify(rects)}`
    const previous = this.owners.get(key)
    if (previous?.signature === signature) return false
    if (previous) this.unindex(key, previous)
    if (rects.length === 0) {
      this.owners.delete(key)
    } else {
      const entry: OwnerEntry = { owner, rects, signature, seq: ++this.seq }
      this.owners.set(key, entry)
      this.index(key, entry)
    }
    this.emit(owner.unitId)
    return true
  }

  /** Format-reading owners register their inputs: Univer skips recalculation for format-only edits. */
  watchSources(owner: FormatOwner, rects: readonly SourceRect[]): void {
    const key = ownerKey(owner)
    if (rects.length === 0) this.sources.delete(key)
    else this.sources.set(key, { unitId: owner.unitId, rects })
  }

  /** True when a format-only edit of `range` must recalculate some owner. */
  isWatched(unitId: string, sheetId: string, range: IRange): boolean {
    for (const watch of this.sources.values()) {
      if (watch.unitId !== unitId) continue
      for (const r of watch.rects) {
        if (r.sheetId !== sheetId) continue
        if (
          range.startRow <= r.endRow &&
          range.endRow >= r.startRow &&
          range.startColumn <= r.endColumn &&
          range.endColumn >= r.startColumn
        )
          return true
      }
    }
    return false
  }

  clear(owner: Pick<FormatOwner, 'unitId' | 'sheetId' | 'row' | 'column'>): void {
    const key = ownerKey(owner)
    this.sources.delete(key)
    const previous = this.owners.get(key)
    if (!previous) return
    this.unindex(key, previous)
    this.owners.delete(key)
    this.emit(owner.unitId)
  }

  clearUnit(unitId: string): void {
    for (const [key, watch] of [...this.sources])
      if (watch.unitId === unitId) this.sources.delete(key)
    for (const [key, entry] of [...this.owners]) {
      if (entry.owner.unitId !== unitId) continue
      this.unindex(key, entry)
      this.owners.delete(key)
    }
  }

  ownersOf(unitId: string): FormatOwner[] {
    return [...this.owners.values()].filter((e) => e.owner.unitId === unitId).map((e) => e.owner)
  }

  /** Merged format patch for one cell, or undefined when no DVH function formats it. */
  styleAt(
    unitId: string,
    sheetId: string,
    row: number,
    column: number,
  ): WorkbookStyleEdit | undefined {
    const sk = sheetKey(unitId, sheetId)
    const hits: { seq: number; style: WorkbookStyleEdit }[] = []
    for (const key of this.points.get(sk)?.get(cellKey(row, column)) ?? []) {
      const entry = this.owners.get(key)
      const rect = entry?.rects.find(
        (r) => r.sheetId === sheetId && r.startRow === row && r.startColumn === column,
      )
      if (entry && rect) hits.push({ seq: entry.seq, style: rect.style })
    }
    for (const key of this.areas.get(sk) ?? []) {
      const entry = this.owners.get(key)
      if (!entry) continue
      for (const r of entry.rects) {
        if (r.sheetId !== sheetId || isPoint(r)) continue
        if (
          row >= r.startRow &&
          row <= r.endRow &&
          column >= r.startColumn &&
          column <= r.endColumn
        ) {
          hits.push({ seq: entry.seq, style: r.style })
        }
      }
    }
    if (hits.length === 0) return undefined
    if (hits.length === 1) return hits[0]!.style
    hits.sort((a, b) => a.seq - b.seq)
    return Object.assign({}, ...hits.map((h) => h.style)) as WorkbookStyleEdit
  }

  /** Every formatted cell of the unit with its merged patch, for the save payload. */
  bake(unitId: string): BakedStyle[] {
    const cells = new Map<string, BakedStyle>()
    for (const entry of this.owners.values()) {
      if (entry.owner.unitId !== unitId) continue
      for (const r of entry.rects) {
        let budget = MAX_BAKED_CELLS_PER_AREA
        for (let row = r.startRow; row <= r.endRow && budget > 0; row++) {
          for (
            let column = r.startColumn;
            column <= r.endColumn && budget > 0;
            column++, budget--
          ) {
            const key = `${r.sheetId}\u0000${cellKey(row, column)}`
            if (cells.has(key)) continue
            const style = this.styleAt(unitId, r.sheetId, row, column)
            if (style) cells.set(key, { sheetId: r.sheetId, row, column, style })
          }
        }
      }
    }
    return [...cells.values()]
  }

  onChange(listener: (unitId: string) => void): () => void {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }

  private index(key: string, entry: OwnerEntry): void {
    for (const r of entry.rects) {
      const sk = sheetKey(entry.owner.unitId, r.sheetId)
      if (isPoint(r)) {
        const sheet = this.points.get(sk) ?? new Map<string, Set<string>>()
        this.points.set(sk, sheet)
        const ck = cellKey(r.startRow, r.startColumn)
        const set = sheet.get(ck) ?? new Set<string>()
        sheet.set(ck, set.add(key))
      } else {
        const set = this.areas.get(sk) ?? new Set<string>()
        this.areas.set(sk, set.add(key))
      }
    }
  }

  private unindex(key: string, entry: OwnerEntry): void {
    for (const r of entry.rects) {
      const sk = sheetKey(entry.owner.unitId, r.sheetId)
      if (isPoint(r)) {
        const sheet = this.points.get(sk)
        const ck = cellKey(r.startRow, r.startColumn)
        sheet?.get(ck)?.delete(key)
        if (sheet?.get(ck)?.size === 0) sheet.delete(ck)
      } else {
        this.areas.get(sk)?.delete(key)
      }
    }
  }

  private emit(unitId: string): void {
    for (const listener of this.listeners) listener(unitId)
  }
}

function isPoint(r: FormatRect): boolean {
  return r.startRow === r.endRow && r.startColumn === r.endColumn
}

/** The app-wide store the DVH functions publish into. */
export const formulaFormats = new FormulaFormatStore()

/** Univer style patch for a neutral style; null values remove the attribute. */
const univerPatchCache = new WeakMap<WorkbookStyleEdit, Record<string, unknown>>()
function univerPatch(style: WorkbookStyleEdit): Record<string, unknown> {
  let patch = univerPatchCache.get(style)
  if (!patch) {
    patch = fromNeutralStyle(style)
    univerPatchCache.set(style, patch)
  }
  return patch
}

export function composeFormattedStyle(
  base: Nullable<IStyleData>,
  style: WorkbookStyleEdit,
): IStyleData {
  const merged: Record<string, unknown> = { ...(base ?? {}) }
  for (const [key, value] of Object.entries(univerPatch(style))) {
    if (value === null) delete merged[key]
    else if (key === 'bd' && typeof value === 'object') {
      const borders: Record<string, unknown> = { ...((merged.bd as Record<string, unknown>) ?? {}) }
      for (const [edge, edgeValue] of Object.entries(value as Record<string, unknown>)) {
        if (edgeValue === null) delete borders[edge]
        else borders[edge] = edgeValue
      }
      merged.bd = borders
    } else merged[key] = value
  }
  return merged as IStyleData
}

/** True while the owner cell still holds a formula calling the owner's function. */
export function ownerIsAlive(workbook: Workbook, owner: FormatOwner): boolean {
  const sheet = workbook.getSheetBySheetId(owner.sheetId)
  if (!sheet) return false
  const cell = sheet.getCellRaw(owner.row, owner.column)
  if (!cell) return false
  // a shared formula's dependents carry only `si`; trust them until their own recalc
  if (typeof cell.f !== 'string') return cell.si != null
  return cell.f.toUpperCase().includes(`${owner.functionName}(`)
}

/**
 * Dirties the formula engine for format-only edits of watched ranges. Univer
 * ignores set-range-values from the style/border/clear-format commands (as
 * Excel does); registered with the active dirty manager, this mutation rides
 * Univer's own calculation queue.
 */
export const DVH_FORMAT_SOURCE_CHANGED = 'dvh.mutation.format-source-changed'
const FORMAT_ONLY_TRIGGERS = new Set<string>([
  SetStyleCommand.id,
  SetBorderCommand.id,
  ClearSelectionFormatCommand.id,
])

const OWNER_INVALIDATING_MUTATIONS = new Set<string>([
  SetRangeValuesMutation.id,
  InsertRowMutation.id,
  RemoveRowMutation.id,
  InsertColMutation.id,
  RemoveColMutation.id,
  MoveRowsMutation.id,
  MoveColsMutation.id,
  MoveRangeMutation.id,
  RemoveSheetMutation.id,
])

export function installFormulaFormatChannel(
  runtime: UniverRuntime,
  store: FormulaFormatStore = formulaFormats,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const interceptor = injector
    .get(SheetInterceptorService)
    .intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
      // Above NUMFMT (10): a number format a function returns must reach the
      // number-format interceptor, which reads the cell's style.
      priority: 11,
      effect: InterceptorEffectEnum.Style | InterceptorEffectEnum.Value,
      handler: (cell, location, next) => {
        const style = store.styleAt(location.unitId, location.subUnitId, location.row, location.col)
        if (!style) return next(cell)
        const base = location.workbook.getStyles().getStyleByCell(cell)
        return next({ ...(cell ?? {}), s: composeFormattedStyle(base, style) } as ICellData)
      },
    })

  // Repaint once per frame after formats change; skeletons cache cell styles.
  const pending = new Set<string>()
  let frame: ReturnType<typeof setTimeout> | null = null
  const repaint = (): void => {
    frame = null
    const renders = injector.get(IRenderManagerService)
    for (const unitId of pending) {
      const render = renders.getRenderById(unitId)
      render?.with(SheetSkeletonManagerService).reCalculate()
      render?.mainComponent?.makeDirty()
    }
    pending.clear()
  }
  const stopListening = store.onChange((unitId) => {
    pending.add(unitId)
    frame ??= setTimeout(repaint, 0)
  })

  // A deleted or overwritten formula withdraws its formats; structural edits
  // may move the owner, whose recalc republishes under the new address.
  const instances = injector.get(IUniverInstanceService)
  const commands = injector.get(ICommandService)
  const sourceChangedRegistration = commands.registerCommand({
    id: DVH_FORMAT_SOURCE_CHANGED,
    type: CommandType.MUTATION,
    handler: () => true,
  })
  const dirtyManager = injector.get(IActiveDirtyManagerService)
  dirtyManager.register(DVH_FORMAT_SOURCE_CHANGED, {
    commandId: DVH_FORMAT_SOURCE_CHANGED,
    getDirtyData: (command) => ({
      dirtyRanges:
        (command.params as { dirtyRanges?: IUnitRange[] } | undefined)?.dirtyRanges ?? [],
    }),
  })
  const unitDisposed = instances.unitDisposed$.subscribe((unit) =>
    store.clearUnit(unit.getUnitId()),
  )
  const commandListener = commands.onCommandExecuted((command) => {
    if (!OWNER_INVALIDATING_MUTATIONS.has(command.id)) return
    if (command.id === SetRangeValuesMutation.id) {
      const p = command.params as {
        unitId: string
        subUnitId: string
        trigger?: string
        cellValue?: IObjectMatrixPrimitiveType<Nullable<ICellData>>
      }
      if (p.trigger && FORMAT_ONLY_TRIGGERS.has(p.trigger) && p.cellValue) {
        const dirtyRanges = new ObjectMatrix(p.cellValue)
          .getDiscreteRanges()
          .filter((range) => store.isWatched(p.unitId, p.subUnitId, range))
          .map((range): IUnitRange => ({ unitId: p.unitId, sheetId: p.subUnitId, range }))
        if (dirtyRanges.length > 0) {
          setTimeout(() => {
            void commands.executeCommand(
              DVH_FORMAT_SOURCE_CHANGED,
              { dirtyRanges },
              { onlyLocal: true },
            )
          }, 0)
        }
      }
    }
    const unitId = (command.params as { unitId?: string } | undefined)?.unitId
    if (!unitId) return
    const workbook = instances.getUnit<Workbook>(unitId)
    for (const owner of store.ownersOf(unitId)) {
      if (!workbook || !ownerIsAlive(workbook, owner)) store.clear(owner)
    }
  })

  return {
    dispose() {
      interceptor.dispose()
      stopListening()
      commandListener.dispose()
      sourceChangedRegistration.dispose()
      dirtyManager.remove(DVH_FORMAT_SOURCE_CHANGED)
      unitDisposed.unsubscribe()
      if (frame) clearTimeout(frame)
    },
  }
}

/**
 * Bakes the formats DVH functions currently show into the save payload as
 * style-only edits, on top of the user's own edit of the same cell (the format
 * wins, as it does on screen). The journal and the undo stack never see them.
 */
export function withBakedFormulaFormats(
  edits: readonly WorkbookCellEdit[],
  baked: readonly BakedStyle[],
): WorkbookCellEdit[] {
  if (baked.length === 0) return [...edits]
  const out = [...edits]
  const index = new Map<string, number>()
  out.forEach((edit, i) => index.set(`${edit.sheetId}\u0000${edit.row}:${edit.column}`, i))
  for (const b of baked) {
    const key = `${b.sheetId}\u0000${b.row}:${b.column}`
    const at = index.get(key)
    if (at === undefined) {
      index.set(key, out.length)
      out.push({
        sheetId: b.sheetId,
        row: b.row,
        column: b.column,
        writeValue: false,
        value: null,
        style: b.style,
      })
    } else {
      const edit = out[at]!
      out[at] = { ...edit, style: { ...(edit.style ?? {}), ...b.style } }
    }
  }
  return out
}

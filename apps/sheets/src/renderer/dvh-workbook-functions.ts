// DVH functions that read the workbook beyond their arguments' values: cell
// colors, notes, defined names, formulas and the sheet list. Ported from
// DVH-Tool/Function/clsFun_Color.cs (SumColor, CountColor, ProductColor,
// FilterColor), clsFun_Comment.cs (GetTextCom), clsFun_Name.cs (CountNames),
// clsFun_Explain.cs and clsFun_FilterSheet.cs. They never write the workbook:
// FilterSheet spills its list instead of overwriting the cells below, and the
// tab colors ride on the format channel.
//
// Univer does not recalculate on a format-only edit; the color functions
// register their ranges with the format channel, which re-dirties them when a
// fill or font color changes there. Notes and sheet renames do not trigger a
// recalculation (as in Excel, where these functions are not volatile).
import {
  ArrayValueObject,
  BaseFunction,
  FunctionType,
  IDefinedNamesService,
  IFunctionService,
  NumberValueObject,
  StringValueObject,
  type BaseReferenceObject,
  type BaseValueObject,
  type IFunctionInfo,
} from '@univerjs/engine-formula'
import type { IStyleData, Nullable } from '@univerjs/core'
import { IDescriptionService } from '@univerjs/sheets-formula'
import { isDvhDefinedName } from '@genoffice/dvh-model'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'

import { formulaFormats, type FormatRect, type FormulaFormatStore } from './dvh-format-channel'
import {
  DvhFormattedFunction,
  asReference,
  boundsOf,
  isBlankArg,
  isFalseLike,
  scalarOf,
} from './dvh-format-functions'
import { registerDvhAliases } from './dvh-function-aliases'
import type { UniverRuntime } from './univer-state'

type Scalar = string | number | boolean

/** Cells a color function scans at most (Excel-DNA iterated every cell over COM). */
const MAX_SCANNED_CELLS = 500_000
const MAX_EXPLAINED_CELLS = 10_000

// ---------------- pure helpers ----------------

/** Excel's Interior.Color of a cell without fill, and Font.Color of the default font. */
export const OLE_NO_FILL = 16_777_215
export const OLE_AUTOMATIC_FONT = 0

/** `#RRGGBB`, `#RGB` or `rgb(r, g, b)` as Excel's OLE color (R + G·256 + B·65536); null when unparseable. */
export function cssColorToOle(color: string | null | undefined): number | null {
  if (!color) return null
  const text = color.trim()
  let rgb: [number, number, number] | null = null
  const long = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i.exec(text)
  const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(text)
  const fn = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(text)
  if (long) rgb = [parseInt(long[1]!, 16), parseInt(long[2]!, 16), parseInt(long[3]!, 16)]
  else if (short)
    rgb = [
      parseInt(short[1]! + short[1]!, 16),
      parseInt(short[2]! + short[2]!, 16),
      parseInt(short[3]! + short[3]!, 16),
    ]
  else if (fn)
    rgb = [Number(fn[1]), Number(fn[2]), Number(fn[3])].map((n) => Math.min(255, n)) as [
      number,
      number,
      number,
    ]
  if (!rgb) return null
  return rgb[0] + rgb[1] * 256 + rgb[2] * 65_536
}

/** The add-in's IsNumericValueFixed: numbers and numeric text (blank text is not numeric). */
export function numericValue(value: Scalar | null | undefined): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value !== 'string' || value.trim() === '') return null
  const n = Number(value.trim())
  return Number.isFinite(n) ? n : null
}

/** C#'s Math.Round(value, digits).ToString(): no trailing zeros. */
export function roundText(value: number, digits: number): string {
  const rounded = Number(value.toFixed(Math.max(0, Math.min(15, digits))))
  return String(Object.is(rounded, -0) ? 0 : rounded)
}

/** The add-in's EvaluateCondition for SUMIF criteria (">=5", "<>a", "x"). */
export function matchesCriterion(value: Scalar | null, criterion: string): boolean {
  if (value === null || value === '' || criterion === '') return false
  let op = '='
  let operand = criterion.trim()
  for (const candidate of ['>=', '<=', '<>', '>', '<', '=']) {
    if (operand.startsWith(candidate)) {
      op = candidate
      operand = operand.slice(candidate.length)
      break
    }
  }
  operand = operand.trim().replace(/^"+|"+$/g, '')
  const a = numericValue(value)
  const b = numericValue(operand)
  if (a !== null && b !== null) {
    switch (op) {
      case '=':
        return Math.abs(a - b) < 1e-7
      case '>':
        return a > b
      case '<':
        return a < b
      case '>=':
        return a >= b
      case '<=':
        return a <= b
      default:
        return Math.abs(a - b) > 1e-7
    }
  }
  const equal = String(value).toLowerCase() === operand.toLowerCase()
  return op === '=' ? equal : op === '<>' ? !equal : false
}

/** Splits function arguments on top-level `,`/`;` (quoted text and nested parentheses stay whole). */
export function splitArguments(text: string): string[] {
  const parts: string[] = []
  let depth = 0
  let quoted = false
  let start = 0
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (ch === '"') quoted = !quoted
    else if (quoted) continue
    else if (ch === '(') depth++
    else if (ch === ')') depth--
    else if ((ch === ',' || ch === ';') && depth === 0) {
      parts.push(text.slice(start, i))
      start = i + 1
    }
  }
  parts.push(text.slice(start))
  return parts.map((part) => part.trim())
}

/** Reads the values of an A1 reference ("B3", "B3:C9", "Sheet2!B3"); null when it is not one. */
export type ReferenceReader = (reference: string) => (Scalar | null)[] | null

const REFERENCE_TOKEN =
  /((?:'[^']+'|[A-Za-z_À-￿][\w.À-￿]*)!)?\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?(?![\w(!])/g

/**
 * DVH.Explain: the formula with its references replaced by their values.
 * SUM/PRODUCT spell out every number of their ranges, SUMIF only the summed
 * cells that meet the criterion; anything else swaps each single-cell
 * reference outside quoted text for its (rounded) number.
 */
export function explainFormula(formula: string, digits: number, read: ReferenceReader): string {
  const body = formula.replace(/^=/, '').trim()
  const whole = /^(SUM|PRODUCT|SUMIF)\((.*)\)$/i.exec(body)
  const numbers = (reference: string): string[] =>
    (read(reference.replace(/\$/g, '')) ?? [])
      .map(numericValue)
      .filter((n): n is number => n !== null)
      .map((n) => roundText(n, digits))
  if (whole && findClosing(body, whole[1]!.length) === body.length - 1) {
    const name = whole[1]!.toUpperCase()
    const args = splitArguments(whole[2]!)
    if (name === 'SUMIF') {
      const criteriaRange = read(args[0]?.replace(/\$/g, '') ?? '') ?? []
      const criterion = (args[1] ?? '').replace(/^"|"$/g, '')
      const sumRange = args[2] ? (read(args[2].replace(/\$/g, '')) ?? []) : criteriaRange
      const picked: string[] = []
      criteriaRange.forEach((value, i) => {
        if (!matchesCriterion(value, criterion)) return
        const n = numericValue(sumRange[i] ?? null)
        if (n !== null) picked.push(roundText(n, digits))
      })
      return picked.length === 0 ? '=0' : `=${picked.join('+')}`
    }
    const terms = args.flatMap((arg) => numbers(arg))
    return `=${terms.join(name === 'SUM' ? '+' : '*')}`
  }
  // Normal formulas: swap single-cell references outside quoted text.
  let result = ''
  const segments = body.split('"')
  segments.forEach((segment, index) => {
    if (index % 2 === 1) {
      result += `"${segment}"`
      return
    }
    result += segment.replace(REFERENCE_TOKEN, (token: string) => {
      if (token.includes(':')) return token.replace(/\$/g, '')
      const values = read(token.replace(/\$/g, ''))
      const n = values && values.length === 1 ? numericValue(values[0] ?? null) : null
      return n === null ? token.replace(/\$/g, '') : roundText(n, digits)
    })
  })
  return `=${result}`
}

function findClosing(text: string, open: number): number {
  let depth = 0
  let quoted = false
  for (let i = open; i < text.length; i++) {
    const ch = text[i]
    if (ch === '"') quoted = !quoted
    else if (quoted) continue
    else if (ch === '(') depth++
    else if (ch === ')' && --depth === 0) return i
  }
  return -1
}

/** Explains one static value the way the add-in did. */
export function explainValue(
  value: Scalar | null,
  digits: number,
  numberLabel = 'Giá trị',
): string {
  if (value === null || value === '') return '(Trống)'
  const n = numericValue(value)
  if (n !== null) return `${numberLabel}: ${roundText(n, digits)}`
  return `Nội dung: ${String(value)}`
}

export interface NameInfo {
  readonly name: string
  readonly ref: string
  readonly hidden: boolean
}

const NAME_ERRORS = ['#REF!', '#NAME?', '#VALUE!', '#N/A', '#NULL!', '#DIV/0!']

/** clsFun_Name's categories: ALL, ERR (error refs), OUT (external links), HIDDEN. */
export function nameMatchesCategory(info: NameInfo, category: string): boolean {
  switch (category.toUpperCase().trim()) {
    case 'ALL':
      return true
    case 'ERR':
      return (
        info.ref.trim() === '' ||
        NAME_ERRORS.some((error) => info.ref.toUpperCase().includes(error))
      )
    case 'OUT':
      return (
        (info.ref.includes('[') && info.ref.includes(']')) ||
        info.ref.includes(':\\') ||
        info.ref.includes('\\\\') ||
        /https?:\/\//i.test(info.ref)
      )
    case 'HIDDEN':
      return info.hidden
    default:
      return false
  }
}

/** DVH.FilterSheet's list: a header, then every sheet whose name contains `text` (case-insensitive). */
export function filterSheetNames(names: readonly string[], text: string): string[] {
  const search = text.trim().toLowerCase()
  const header = search === '' ? 'List of all sheets:' : `Sheets matching '${search}':`
  return [header, ...names.filter((name) => search === '' || name.toLowerCase().includes(search))]
}

// ---------------- cell colors ----------------

function cellOle(
  styles: { getStyleByCell(cell: unknown): Nullable<IStyleData> },
  reference: BaseReferenceObject,
  store: FormulaFormatStore,
  unitId: string,
  row: number,
  column: number,
  font: boolean,
): number {
  const sheetId = reference.getSheetId()
  const overlay = store.styleAt(unitId, sheetId, row, column)
  const style = styles.getStyleByCell(reference.getCellData(row, column))
  if (font) {
    const color = overlay?.fontColor !== undefined ? overlay.fontColor : style?.cl?.rgb
    return cssColorToOle(typeof color === 'string' ? color : null) ?? OLE_AUTOMATIC_FONT
  }
  const fill = overlay?.fillColor !== undefined ? overlay.fillColor : style?.bg?.rgb
  return cssColorToOle(typeof fill === 'string' ? fill : null) ?? OLE_NO_FILL
}

abstract class DvhColorReader extends DvhFormattedFunction {
  /** The color to match: a number is an OLE color, a reference lends its first cell's color. */
  protected targetColor(value: BaseValueObject | undefined, font: boolean): number | null {
    if (isBlankArg(value)) return null
    if (value.isReferenceObject()) {
      const reference = asReference(value)
      const bounds = boundsOf(reference)
      this.watch(reference)
      return cellOle(
        reference.getCurrentStylesData(),
        reference,
        this.store,
        this.unitId ?? '',
        bounds.startRow,
        bounds.startColumn,
        font,
      )
    }
    const n = numericValue(scalarOf(value))
    return n === null ? null : Math.trunc(n)
  }
  private watched: {
    sheetId: string
    startRow: number
    endRow: number
    startColumn: number
    endColumn: number
  }[] = []
  protected watch(reference: BaseReferenceObject): void {
    this.watched.push({ sheetId: reference.getSheetId(), ...boundsOf(reference) })
  }
  /** Format-only edits of the read ranges must recalc this cell (see dvh-format-channel). */
  protected commitWatches(): void {
    const owner = this.owner()
    // splice: an alias (dvh-function-aliases.ts) shares this array through its prototype
    const watched = this.watched.splice(0)
    if (owner) this.store.watchSources(owner, watched)
  }
}

class DvhColorAggregate extends DvhColorReader {
  override minParams = 2
  override maxParams = 3
  constructor(
    private readonly mode: 'sum' | 'count' | 'product',
    store: FormulaFormatStore,
  ) {
    super(
      mode === 'sum' ? 'DVH.SumColor' : mode === 'count' ? 'DVH.CountColor' : 'DVH.ProductColor',
      store,
    )
  }
  override calculate(
    range: BaseValueObject,
    color: BaseValueObject,
    isFont?: BaseValueObject,
  ): BaseValueObject {
    try {
      if (!range.isReferenceObject()) return StringValueObject.create('#REF: Invalid range')
      const font = !isFalseLike(scalarOf(isFont))
      const target = this.targetColor(color, font)
      if (target === null) return StringValueObject.create('#VALUE: Invalid color')
      const reference = asReference(range)
      this.watch(reference)
      const bounds = boundsOf(reference)
      const rows = bounds.endRow - bounds.startRow + 1
      const columns = bounds.endColumn - bounds.startColumn + 1
      if (rows * columns > MAX_SCANNED_CELLS)
        return StringValueObject.create('#VALUE: Range too large')
      const values = reference.toArrayValueObject()
      const styles = reference.getCurrentStylesData()
      const matched: number[] = []
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < columns; c++) {
          const cell = values.get(r, c)
          if (!cell || cell.isNull() || cell.isError()) continue
          const n = numericValue(cell.getValue() as Scalar)
          if (n === null) continue
          if (this.mode === 'sum' && n <= 0) continue
          const ole = cellOle(
            styles,
            reference,
            this.store,
            this.unitId ?? '',
            bounds.startRow + r,
            bounds.startColumn + c,
            font,
          )
          if (ole === target) matched.push(n)
        }
      }
      if (this.mode === 'count') return NumberValueObject.create(matched.length)
      if (this.mode === 'sum') return NumberValueObject.create(matched.reduce((a, b) => a + b, 0))
      if (matched.length === 0 || matched.some((n) => Math.abs(n) < Number.EPSILON))
        return NumberValueObject.create(0)
      let product = 1
      for (const n of matched) {
        product *= n
        if (!Number.isFinite(product)) return StringValueObject.create('#NUM: Overflow')
      }
      return NumberValueObject.create(product)
    } finally {
      this.commitWatches()
    }
  }
}

class DvhFilterColor extends DvhColorReader {
  override minParams = 3
  override maxParams = 4
  constructor(store: FormulaFormatStore) {
    super('DVH.FilterColor', store)
  }
  override calculate(
    range: BaseValueObject,
    columnIndex: BaseValueObject,
    color: BaseValueObject,
    isFont?: BaseValueObject,
  ): BaseValueObject {
    try {
      if (!range.isReferenceObject()) return StringValueObject.create('#REF: Invalid range')
      const reference = asReference(range)
      this.watch(reference)
      const bounds = boundsOf(reference)
      const rows = bounds.endRow - bounds.startRow + 1
      const columns = bounds.endColumn - bounds.startColumn + 1
      if (rows * columns > MAX_SCANNED_CELLS)
        return StringValueObject.create('#VALUE: Range too large')
      const index = Math.trunc(numericValue(scalarOf(columnIndex)) ?? 0)
      if (index < 1 || index > columns)
        return StringValueObject.create(`#COLUMN: Must be 1-${columns}`)
      const font = !isFalseLike(scalarOf(isFont))
      const target = this.targetColor(color, font)
      if (target === null) return StringValueObject.create('#VALUE: Invalid color')
      const values = reference.toArrayValueObject()
      const styles = reference.getCurrentStylesData()
      const kept: BaseValueObject[][] = []
      for (let r = 0; r < rows; r++) {
        const ole = cellOle(
          styles,
          reference,
          this.store,
          this.unitId ?? '',
          bounds.startRow + r,
          bounds.startColumn + index - 1,
          font,
        )
        if (ole !== target) continue
        kept.push(
          Array.from({ length: columns }, (_, c) => {
            const cell = values.get(r, c)
            return cell == null || cell.isNull() ? StringValueObject.create('') : cell
          }),
        )
      }
      if (kept.length === 0) return StringValueObject.create('No matches found')
      return ArrayValueObject.create({
        calculateValueList: kept,
        rowCount: kept.length,
        columnCount: columns,
        unitId: this.unitId ?? '',
        sheetId: this.subUnitId ?? '',
        row: this.row,
        column: this.column,
      })
    } finally {
      this.commitWatches()
    }
  }
}

// ---------------- notes, names, formulas, sheets ----------------

abstract class DvhWorkbookReader extends BaseFunction {
  override needsReferenceObject = true
  constructor(
    name: string,
    protected readonly runtime: UniverRuntime,
  ) {
    super(name.toUpperCase())
  }
  protected worksheet(sheetId: string) {
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    if (!workbook || (this.unitId && workbook.getId() !== this.unitId)) return null
    return workbook.getSheetBySheetId(sheetId)
  }
}

class DvhGetTextCom extends DvhWorkbookReader {
  override minParams = 1
  override maxParams = 3
  constructor(runtime: UniverRuntime) {
    super('DVH.GetTextCom', runtime)
  }
  override calculate(
    range: BaseValueObject,
    delimiter?: BaseValueObject,
    includeAddress?: BaseValueObject,
  ): BaseValueObject {
    if (!range.isReferenceObject()) return StringValueObject.create('#ERROR: Invalid range')
    const reference = asReference(range)
    const bounds = boundsOf(reference)
    const sheet = this.worksheet(reference.getSheetId())
    if (!sheet) return StringValueObject.create('#ERROR: Invalid range')
    const separator = isBlankArg(delimiter) ? ', ' : String(scalarOf(delimiter) ?? '')
    const withAddress = !isFalseLike(scalarOf(includeAddress))
    const notes = (sheet as unknown as { getNotes(): { row: number; col: number; note: string }[] })
      .getNotes()
      .filter(
        (n) =>
          n.row >= bounds.startRow &&
          n.row <= bounds.endRow &&
          n.col >= bounds.startColumn &&
          n.col <= bounds.endColumn &&
          n.note.trim() !== '',
      )
      .sort((a, b) => a.row - b.row || a.col - b.col)
    if (notes.length === 0)
      return StringValueObject.create('No comments found / Không tìm thấy chú thích')
    return StringValueObject.create(
      notes
        .map((n) => (withAddress ? `${columnLabel(n.col)}${n.row + 1}: ${n.note}` : n.note))
        .join(separator),
    )
  }
}

class DvhCountNames extends DvhWorkbookReader {
  override minParams = 0
  override maxParams = 1
  constructor(runtime: UniverRuntime) {
    super('DVH.CountNames', runtime)
  }
  override calculate(type?: BaseValueObject): BaseValueObject {
    const category = isBlankArg(type) ? 'ALL' : String(scalarOf(type) ?? 'ALL')
    const names = workbookNames(this.runtime, this.unitId ?? '')
    return NumberValueObject.create(
      names.filter((info) => nameMatchesCategory(info, category)).length,
    )
  }
}

/** The workbook's defined names, without DVH's own binding names. */
export function workbookNames(runtime: UniverRuntime, unitId: string): NameInfo[] {
  const map =
    runtime.univer.__getInjector().get(IDefinedNamesService).getDefinedNameMap(unitId) ?? {}
  return Object.values(map)
    .filter((item) => !isDvhDefinedName(item.name))
    .map((item) => ({
      name: item.name,
      ref: item.formulaOrRefString ?? '',
      hidden: item.hidden === true,
    }))
}

class DvhExplain extends DvhWorkbookReader {
  override minParams = 1
  override maxParams = 2
  constructor(runtime: UniverRuntime) {
    super('DVH.Explain', runtime)
  }
  override calculate(target: BaseValueObject, digitsArg?: BaseValueObject): BaseValueObject {
    const rawDigits = numericValue(scalarOf(digitsArg))
    const digits = Math.max(0, Math.min(15, Math.trunc(rawDigits ?? 2)))
    if (!target.isReferenceObject()) {
      return StringValueObject.create(explainValue(scalarOf(target), digits, 'Giá trị số'))
    }
    const reference = asReference(target)
    const bounds = boundsOf(reference)
    const sheet = this.worksheet(reference.getSheetId())
    if (!sheet) return StringValueObject.create('#ERROR: Invalid range')
    const rows = bounds.endRow - bounds.startRow + 1
    const columns = bounds.endColumn - bounds.startColumn + 1
    if (rows * columns > MAX_EXPLAINED_CELLS)
      return StringValueObject.create('#ERROR: Range too large')
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    const read: ReferenceReader = (text) => {
      try {
        const bang = text.lastIndexOf('!')
        const target =
          bang < 0
            ? sheet
            : workbook?.getSheetByName(
                text.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'"),
              )
        if (!target) return null
        return target
          .getRange(bang < 0 ? text : text.slice(bang + 1))
          .getValues()
          .flat()
          .map((v) => (v === undefined ? null : (v as Scalar | null)))
      } catch {
        return null
      }
    }
    const explainCell = (row: number, column: number): string => {
      const range = sheet.getRange(row, column, 1, 1)
      const formula = range.getFormula()
      if (formula && formula.startsWith('=')) {
        try {
          return explainFormula(formula, digits, read)
        } catch (error) {
          return `#ERROR: ${error instanceof Error ? error.message : String(error)}`
        }
      }
      const value = range.getValue()
      return explainValue(value === undefined ? null : (value as Scalar | null), digits)
    }
    if (rows === 1 && columns === 1) {
      return StringValueObject.create(explainCell(bounds.startRow, bounds.startColumn))
    }
    return ArrayValueObject.create({
      calculateValueList: Array.from({ length: rows }, (_, r) =>
        Array.from({ length: columns }, (_, c) =>
          StringValueObject.create(explainCell(bounds.startRow + r, bounds.startColumn + c)),
        ),
      ),
      rowCount: rows,
      columnCount: columns,
      unitId: this.unitId ?? '',
      sheetId: this.subUnitId ?? '',
      row: this.row,
      column: this.column,
    })
  }
}

class DvhFilterSheet extends DvhFormattedFunction {
  override minParams = 0
  override maxParams = 1
  constructor(
    private readonly runtime: UniverRuntime,
    store: FormulaFormatStore,
  ) {
    super('DVH.FilterSheet', store)
  }
  override calculate(text?: BaseValueObject): BaseValueObject {
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    if (!workbook || (this.unitId && workbook.getId() !== this.unitId))
      return this.fail('Context Error')
    const sheets = workbook.getSheets()
    const search = isBlankArg(text) ? '' : String(scalarOf(text) ?? '')
    const list = filterSheetNames(
      sheets.map((sheet) => sheet.getSheetName()),
      search,
    )
    const colorOf = new Map(sheets.map((sheet) => [sheet.getSheetName(), sheet.getTabColor()]))
    const sheetId = this.subUnitId ?? ''
    // Tab colors paint the listed names, as the add-in's Interior.Color did.
    const rects: FormatRect[] = []
    list.forEach((name, i) => {
      const color = i === 0 ? undefined : colorOf.get(name)
      if (!color) return
      const row = this.row + i
      rects.push({
        sheetId,
        startRow: row,
        endRow: row,
        startColumn: this.column,
        endColumn: this.column,
        style: { fillColor: color },
      })
    })
    this.publish(rects)
    return ArrayValueObject.create({
      calculateValueList: list.map((name) => [StringValueObject.create(name)]),
      rowCount: list.length,
      columnCount: 1,
      unitId: this.unitId ?? '',
      sheetId,
      row: this.row,
      column: this.column,
    })
  }
}

const descriptions: Pick<IFunctionInfo, 'functionName' | 'abstract' | 'functionParameter'>[] = [
  ...(['SUMCOLOR', 'COUNTCOLOR', 'PRODUCTCOLOR'] as const).map((name) => ({
    functionName: `DVH.${name}`,
    abstract:
      name === 'SUMCOLOR'
        ? 'Tổng các số dương có màu nền/màu chữ trùng màu chỉ định'
        : name === 'COUNTCOLOR'
          ? 'Đếm các ô số có màu nền/màu chữ trùng màu chỉ định'
          : 'Tích các ô số có màu nền/màu chữ trùng màu chỉ định',
    functionParameter: [
      { name: 'range', detail: 'Vùng cần tính', example: 'A1:A100', require: 1, repeat: 0 },
      {
        name: 'color_value',
        detail: 'Ô mẫu màu hoặc mã màu OLE',
        example: 'C1',
        require: 1,
        repeat: 0,
      },
      {
        name: 'is_font_color',
        detail: 'TRUE = màu chữ, FALSE = màu nền (mặc định)',
        example: 'FALSE',
        require: 0,
        repeat: 0,
      },
    ],
  })),
  {
    functionName: 'DVH.FILTERCOLOR',
    abstract: 'Lọc các hàng có ô ở cột chỉ định trùng màu',
    functionParameter: [
      { name: 'range', detail: 'Vùng dữ liệu cần lọc', example: 'A1:F100', require: 1, repeat: 0 },
      {
        name: 'column_index',
        detail: 'Cột kiểm tra màu (từ 1)',
        example: '2',
        require: 1,
        repeat: 0,
      },
      {
        name: 'color_value',
        detail: 'Ô mẫu màu hoặc mã màu OLE',
        example: 'H1',
        require: 1,
        repeat: 0,
      },
      {
        name: 'is_font_color',
        detail: 'TRUE = màu chữ, FALSE = màu nền (mặc định)',
        example: 'FALSE',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.GETTEXTCOM',
    abstract: 'Nối nội dung chú thích (note) trong vùng',
    functionParameter: [
      {
        name: 'source_range',
        detail: 'Vùng chứa chú thích',
        example: 'A1:D20',
        require: 1,
        repeat: 0,
      },
      {
        name: 'delimiter',
        detail: 'Ký tự ngăn cách (mặc định ", ")',
        example: '"; "',
        require: 0,
        repeat: 0,
      },
      {
        name: 'include_address',
        detail: 'TRUE = kèm địa chỉ ô',
        example: 'TRUE',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.COUNTNAMES',
    abstract: 'Đếm Name trong workbook theo loại: ALL, ERR, OUT, HIDDEN',
    functionParameter: [
      { name: 'type', detail: 'ALL, ERR, OUT, HIDDEN', example: '"ERR"', require: 0, repeat: 0 },
    ],
  },
  {
    functionName: 'DVH.EXPLAIN',
    abstract: 'Diễn giải công thức: thay tham chiếu bằng giá trị',
    functionParameter: [
      { name: 'target', detail: 'Ô/vùng cần diễn giải', example: 'E5', require: 1, repeat: 0 },
      {
        name: 'decimal_digits',
        detail: 'Số chữ số thập phân (0-15, mặc định 2)',
        example: '2',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.FILTERSHEET',
    abstract: 'Liệt kê các sheet có tên chứa chuỗi tìm (tô màu theo màu tab)',
    functionParameter: [
      {
        name: 'text',
        detail: 'Chuỗi tìm (để trống: tất cả)',
        example: '"KL"',
        require: 0,
        repeat: 0,
      },
    ],
  },
]

export const DVH_WORKBOOK_FUNCTION_NAMES = descriptions.map((d) => d.functionName)

export function installDvhWorkbookFunctions(
  runtime: UniverRuntime,
  store: FormulaFormatStore = formulaFormats,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const functions = injector.get(IFunctionService)
  const executors: BaseFunction[] = [
    new DvhColorAggregate('sum', store),
    new DvhColorAggregate('count', store),
    new DvhColorAggregate('product', store),
    new DvhFilterColor(store),
    new DvhGetTextCom(runtime),
    new DvhCountNames(runtime),
    new DvhExplain(runtime),
    new DvhFilterSheet(runtime, store),
  ]
  functions.registerExecutors(...executors)
  const infos = descriptions.map((d): IFunctionInfo => ({
    ...d,
    functionType: FunctionType.User,
    description: `DVH Tool · ${d.abstract}`,
  }))
  const descriptionHandle = injector.get(IDescriptionService).registerDescriptions(infos)
  const aliases = registerDvhAliases(injector, executors, infos)
  return {
    dispose() {
      aliases.dispose()
      descriptionHandle.dispose()
      functions.unregisterExecutors(...executors.map((executor) => executor.name))
    },
  }
}

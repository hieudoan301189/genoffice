// DVH functions that return formatting (spike S5). Ported in behavior from
// DVH-Excel/Function/clsFun_Font.cs, clsFun_Color.cs (FillColor, Font.Color)
// and clsFun_Table.cs, but declarative: the formats ride on the format channel
// (dvh-format-channel.ts) instead of being written into the workbook while
// Excel recalculates. Deliberate differences from the add-in:
//  - DVH.Font sets the given styles; the add-in toggled them on every recalc.
//  - DVH.Table spills at the formula cell when range_result is left empty;
//    the legacy form with range_result elsewhere is not computed yet (a file
//    formula keeps the add-in's cached status through the cached-value fallback).
import {
  ArrayValueObject,
  BaseFunction,
  ErrorType,
  ErrorValueObject,
  FunctionType,
  IFunctionService,
  StringValueObject,
  type BaseReferenceObject,
  type BaseValueObject,
  type IFunctionInfo,
} from '@univerjs/engine-formula'
import type { ICellData, IStyleData, Nullable } from '@univerjs/core'
import { IDescriptionService } from '@univerjs/sheets-formula'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'

import {
  formulaFormats,
  type FormatOwner,
  type FormatRect,
  type FormulaFormatStore,
} from './dvh-format-channel'
import { registerDvhAliases } from './dvh-function-aliases'
import { toNeutralStyle } from './edit-journal'
import type { UniverRuntime } from './univer-state'

type Scalar = string | number | boolean

const MAX_TABLE_SOURCE_CELLS = 100_000

// ---------------- pure helpers ----------------

/** B, I, U, U2 (double underline), N/RESET; codes combine ("BU2"). Null when nothing parses. */
export function parseDvhFontCodes(
  input: string,
): { style: WorkbookStyleEdit; labels: string[] } | null {
  let codes = input.toUpperCase().trim()
  if (codes === '') return null
  if (codes === 'N' || codes === 'RESET') {
    return { style: { bold: false, italic: false, underline: false }, labels: ['Bình thường'] }
  }
  const style: Record<string, unknown> = {}
  const labels: string[] = []
  if (codes.includes('U2')) {
    style.underline = true
    style.underlineStyle = 'double'
    labels.push('Gạch chân kép')
    codes = codes.replaceAll('U2', '')
  } else if (codes.includes('U')) {
    style.underline = true
    style.underlineStyle = 'single'
    labels.push('Gạch chân đơn')
  }
  if (codes.includes('B')) {
    style.bold = true
    labels.push('Đậm')
  }
  if (codes.includes('I')) {
    style.italic = true
    labels.push('Nghiêng')
  }
  return labels.length === 0 ? null : { style: style as WorkbookStyleEdit, labels }
}

/** Excel's OLE color (R + G·256 + B·65536) as #RRGGBB; out-of-range values become black, as in the add-in. */
export function oleColorToHex(value: number): string {
  const ole = Number.isFinite(value) && value >= 0 && value <= 0xffffff ? Math.trunc(value) : 0
  const hex = (n: number) => n.toString(16).padStart(2, '0').toUpperCase()
  return `#${hex(ole & 0xff)}${hex((ole >> 8) & 0xff)}${hex((ole >> 16) & 0xff)}`
}

export function a1Address(r: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}): string {
  const start = `${columnLabel(r.startColumn)}${r.startRow + 1}`
  return r.startRow === r.endRow && r.startColumn === r.endColumn
    ? start
    : `${start}:${columnLabel(r.endColumn)}${r.endRow + 1}`
}

/** The add-in's IsFalse: blank, FALSE, 0 and "0"/"FALSE" text reject a row (errors arrive as blank). */
export function isFalseLike(value: Scalar | null | undefined): boolean {
  if (value === null || value === undefined) return true
  if (typeof value === 'boolean') return !value
  if (typeof value === 'number') return Math.abs(value) < Number.EPSILON
  const text = value.trim()
  return text === '' || text === '0' || text.toUpperCase() === 'FALSE'
}

/** title_rows: a count, or (legacy) a title range whose row count is used; a 1×1 text cell counts as 1. */
export function parseTitleRows(value: Scalar | Scalar[][] | null): number | string {
  if (value === null || value === '') return 0
  if (Array.isArray(value)) {
    if (value.length === 1 && value[0]!.length === 1) {
      const cell = value[0]![0]!
      const n = typeof cell === 'number' ? cell : Number.parseInt(String(cell), 10)
      return Number.isFinite(n) ? parseTitleRows(n) : 1
    }
    return value.length
  }
  const n = typeof value === 'number' ? value : Number(value)
  if (!Number.isInteger(n) || n < 0)
    return 'title_rows phải là số nguyên không âm hoặc vùng tiêu đề.'
  return n
}

/**
 * Body-row indexes kept by the filters. A filter is one row/column of truthy
 * values covering every source row (title rows are skipped) or only the body
 * rows, or a single value applied to all rows.
 */
export function selectTableRows(
  totalRows: number,
  titleRows: number,
  filters: readonly (Scalar | Scalar[][] | null)[],
): number[] | string {
  const bodyRows = totalRows - titleRows
  const columns: (Scalar | null)[][] = []
  for (const filter of filters) {
    if (filter === null || filter === '') continue
    if (!Array.isArray(filter)) {
      columns.push(Array.from({ length: bodyRows }, () => filter))
      continue
    }
    const rows = filter.length
    const cols = filter[0]?.length ?? 0
    if (rows === 1 && cols === 1) {
      columns.push(Array.from({ length: bodyRows }, () => filter[0]![0]!))
      continue
    }
    if (rows !== 1 && cols !== 1) return 'Filter phải là một hàng, một cột hoặc một giá trị.'
    const flat = rows === 1 ? filter[0]! : filter.map((row) => row[0]!)
    const offset = flat.length === totalRows ? titleRows : flat.length === bodyRows ? 0 : -1
    if (offset < 0) return 'Số dòng filter không khớp với area_Data.'
    columns.push(flat.slice(offset, offset + bodyRows))
  }
  const kept: number[] = []
  for (let i = 0; i < bodyRows; i++) {
    if (columns.every((column) => !isFalseLike(column[i]))) kept.push(i)
  }
  return kept
}

/** Every attribute a copied cell carries, so the destination's own formats never show through. */
const COPY_RESET: WorkbookStyleEdit = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontColor: null,
  fillColor: null,
  borderTop: null,
  borderBottom: null,
  borderLeft: null,
  borderRight: null,
  wrapText: false,
  horizontalAlignment: 'general',
  numberFormat: 'General',
}

export function copiedCellStyle(source: Nullable<IStyleData>): WorkbookStyleEdit {
  const neutral = source ? toNeutralStyle(source as Record<string, unknown>) : undefined
  return { ...COPY_RESET, ...(neutral ?? {}) }
}

// ---------------- argument access ----------------

/** Univer hands references to `needsReferenceObject` functions as BaseValueObject. */
export const asReference = (value: BaseValueObject): BaseReferenceObject =>
  value as unknown as BaseReferenceObject

/** A missing or empty argument; reference objects have no isNull of their own. */
export function isBlankArg(value: BaseValueObject | undefined): value is undefined {
  return !value || (!value.isReferenceObject() && value.isNull())
}

export function scalarOf(value: BaseValueObject | undefined): Scalar | null {
  if (isBlankArg(value)) return null
  let cell: BaseValueObject = value
  if (cell.isReferenceObject()) cell = asReference(cell).toArrayValueObject()
  if (cell.isArray()) cell = (cell as ArrayValueObject).getFirstCell()
  if (cell.isError()) return null
  return cell.isNull() ? '' : (cell.getValue() as Scalar)
}

export function matrixOf(value: BaseValueObject | undefined): Scalar | Scalar[][] | null {
  if (isBlankArg(value)) return null
  let array: BaseValueObject = value
  if (value.isReferenceObject()) array = asReference(value).toArrayValueObject()
  if (!array.isArray()) return array.isError() ? null : (array.getValue() as Scalar)
  const cells = array as ArrayValueObject
  return Array.from({ length: cells.getRowCount() }, (_, row) =>
    Array.from({ length: cells.getColumnCount() }, (_, column) => {
      const cell = cells.get(row, column)
      return cell == null || cell.isNull() || cell.isError() ? '' : (cell.getValue() as Scalar)
    }),
  )
}

/** A reference's rectangle clamped to the sheet (whole-column refs report open ends). */
export function boundsOf(reference: BaseReferenceObject): {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
} {
  const r = reference.getRangePosition()
  const lastRow = Math.max(0, reference.getActiveSheetRowCount() - 1)
  const lastColumn = Math.max(0, reference.getActiveSheetColumnCount() - 1)
  const clamp = (v: number, fallback: number, max: number) =>
    Number.isFinite(v) && v >= 0 ? Math.min(v, max) : fallback
  return {
    startRow: clamp(r.startRow, 0, lastRow),
    endRow: clamp(r.endRow, lastRow, lastRow),
    startColumn: clamp(r.startColumn, 0, lastColumn),
    endColumn: clamp(r.endColumn, lastColumn, lastColumn),
  }
}

// ---------------- functions ----------------

export abstract class DvhFormattedFunction extends BaseFunction {
  override needsReferenceObject = true
  constructor(
    name: string,
    protected readonly store: FormulaFormatStore,
  ) {
    super(name.toUpperCase())
  }
  protected owner(): FormatOwner | null {
    if (!this.unitId || !this.subUnitId) return null
    return {
      unitId: this.unitId,
      sheetId: this.subUnitId,
      row: this.row,
      column: this.column,
      functionName: String(this.name),
    }
  }
  /** Replaces this formula cell's published formats (an empty list withdraws them). */
  protected publish(rects: FormatRect[]): void {
    const owner = this.owner()
    if (owner) this.store.replace(owner, rects)
  }
  protected fail(message: string): BaseValueObject {
    this.publish([])
    return StringValueObject.create(message)
  }
}

class DvhFontFunction extends DvhFormattedFunction {
  override minParams = 2
  override maxParams = 2
  constructor(store: FormulaFormatStore) {
    super('DVH.Font', store)
  }
  override calculate(range: BaseValueObject, codes: BaseValueObject): BaseValueObject {
    if (!range.isReferenceObject()) return this.fail('#REF: Range không hợp lệ')
    const text = scalarOf(codes)
    if (text === null || String(text).trim() === '') return this.fail('#VALUE: Thiếu mã định dạng')
    const parsed = parseDvhFontCodes(String(text))
    if (!parsed) return this.fail('#VALUE: Mã định dạng không hợp lệ (sử dụng B, I, U, U2, N)')
    const reference = asReference(range)
    const bounds = boundsOf(reference)
    this.publish([{ sheetId: reference.getSheetId(), ...bounds, style: parsed.style }])
    return StringValueObject.create(`✓ ${parsed.labels.join(', ')} → ${a1Address(bounds)}`)
  }
}

class DvhColorFunction extends DvhFormattedFunction {
  override minParams = 1
  override maxParams = 2
  constructor(
    private readonly target: 'fill' | 'font',
    store: FormulaFormatStore,
  ) {
    super(target === 'fill' ? 'DVH.FillColor' : 'DVH.Font.Color', store)
  }
  override calculate(range: BaseValueObject, color?: BaseValueObject): BaseValueObject {
    if (!range.isReferenceObject()) return this.fail('#REF: Invalid range')
    const raw = isBlankArg(color) ? 255 : Number(scalarOf(color))
    const ole = Number.isFinite(raw) && raw >= 0 && raw <= 0xffffff ? Math.trunc(raw) : 0
    const hex = oleColorToHex(ole)
    const reference = asReference(range)
    const bounds = boundsOf(reference)
    const style: WorkbookStyleEdit =
      this.target === 'fill' ? { fillColor: hex } : { fontColor: hex }
    this.publish([{ sheetId: reference.getSheetId(), ...bounds, style }])
    return StringValueObject.create(`Đã đổi màu vùng ${a1Address(bounds)} thành mã màu: ${ole}`)
  }
}

class DvhTableFunction extends DvhFormattedFunction {
  override minParams = 2
  override maxParams = 255
  constructor(store: FormulaFormatStore) {
    super('DVH.Table', store)
  }
  override calculate(
    area: BaseValueObject,
    titleRows: BaseValueObject,
    rangeResult?: BaseValueObject,
    ...filters: BaseValueObject[]
  ): BaseValueObject {
    if (!area.isReferenceObject()) {
      this.publish([])
      return ErrorValueObject.create(ErrorType.VALUE)
    }
    if (rangeResult?.isReferenceObject()) {
      // Legacy target form: writing cells elsewhere is a command, not a formula result.
      this.publish([])
      return ErrorValueObject.create(ErrorType.NA)
    }
    const source = asReference(area)
    const bounds = source.getRangePosition()
    // source formats feed the result: a format-only edit there must recalc this cell
    const owner = this.owner()
    if (owner) this.store.watchSources(owner, [{ sheetId: source.getSheetId(), ...bounds }])
    const totalRows = bounds.endRow - bounds.startRow + 1
    const totalColumns = bounds.endColumn - bounds.startColumn + 1
    if (!(totalRows > 0 && totalColumns > 0) || totalRows * totalColumns > MAX_TABLE_SOURCE_CELLS) {
      return this.fail('❌ area_Data quá lớn hoặc không hợp lệ.')
    }
    const titles = parseTitleRows(matrixOf(titleRows))
    if (typeof titles === 'string') return this.fail(`❌ ${titles}`)
    if (titles >= totalRows)
      return this.fail(`❌ title_rows (${titles}) phải từ 0 đến ${totalRows - 1}.`)
    const kept = selectTableRows(totalRows, titles, filters.map(matrixOf))
    if (typeof kept === 'string') return this.fail(`❌ ${kept}`)

    const sourceRows = [
      ...Array.from({ length: titles }, (_, i) => i),
      ...kept.map((i) => titles + i),
    ]
    if (sourceRows.length === 0) {
      this.publish([])
      return StringValueObject.create('')
    }
    const values = source.toArrayValueObject()
    const styles = source.getCurrentStylesData()
    const sheetId = this.subUnitId ?? ''
    const rects: FormatRect[] = []
    const calculateValueList = sourceRows.map((sourceRow, i) =>
      Array.from({ length: totalColumns }, (_, j) => {
        const raw: Nullable<ICellData> = source.getCellData(
          bounds.startRow + sourceRow,
          bounds.startColumn + j,
        )
        const style = copiedCellStyle(styles.getStyleByCell(raw))
        const row = this.row + i
        const column = this.column + j
        rects.push({
          sheetId,
          startRow: row,
          endRow: row,
          startColumn: column,
          endColumn: column,
          style,
        })
        return values.get(sourceRow, j) ?? null
      }),
    )
    this.publish(rects)
    return ArrayValueObject.create({
      calculateValueList,
      rowCount: sourceRows.length,
      columnCount: totalColumns,
      unitId: this.unitId ?? '',
      sheetId,
      row: this.row,
      column: this.column,
    })
  }
}

const descriptions: Pick<IFunctionInfo, 'functionName' | 'abstract' | 'functionParameter'>[] = [
  {
    functionName: 'DVH.FONT',
    abstract: 'Định dạng chữ cho vùng: B (đậm), I (nghiêng), U, U2 (gạch chân), N (bình thường)',
    functionParameter: [
      { name: 'range', detail: 'Vùng cần định dạng', example: 'A1:A10', require: 1, repeat: 0 },
      {
        name: 'format_style',
        detail: 'Mã kiểu: B, I, U, U2, N (kết hợp được: BI, BU2)',
        example: '"BI"',
        require: 1,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.FILLCOLOR',
    abstract: 'Tô màu nền vùng ô theo mã màu OLE',
    functionParameter: [
      { name: 'range', detail: 'Vùng cần tô màu nền', example: 'A1:B2', require: 1, repeat: 0 },
      {
        name: 'color_value',
        detail: 'Mã màu OLE (mặc định 255 = đỏ)',
        example: '65535',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.FONT.COLOR',
    abstract: 'Đổi màu chữ của vùng ô theo mã màu OLE',
    functionParameter: [
      { name: 'range', detail: 'Vùng cần đổi màu chữ', example: 'A1:B2', require: 1, repeat: 0 },
      {
        name: 'color_value',
        detail: 'Mã màu OLE (mặc định 255 = đỏ)',
        example: '255',
        require: 0,
        repeat: 0,
      },
    ],
  },
  {
    functionName: 'DVH.TABLE',
    abstract:
      'Lọc vùng dữ liệu, giữ định dạng nguồn; để trống range_result để kết quả tràn tại ô công thức',
    functionParameter: [
      {
        name: 'area_Data',
        detail: 'Vùng dữ liệu nguồn (gồm cả hàng tiêu đề)',
        example: 'A1:F100',
        require: 1,
        repeat: 0,
      },
      {
        name: 'title_rows',
        detail: 'Số hàng tiêu đề hoặc vùng tiêu đề',
        example: '2',
        require: 1,
        repeat: 0,
      },
      {
        name: 'range_result',
        detail: 'Để trống: kết quả tràn tại ô công thức',
        example: '',
        require: 0,
        repeat: 0,
      },
      {
        name: 'filters',
        detail: 'Điều kiện lọc: một cột/hàng TRUE/FALSE hoặc một giá trị',
        example: 'C1:C100="Đạt"',
        require: 0,
        repeat: 1,
      },
    ],
  },
]

export const DVH_FORMAT_FUNCTION_NAMES = descriptions.map((d) => d.functionName)

export function installDvhFormatFunctions(
  runtime: UniverRuntime,
  store: FormulaFormatStore = formulaFormats,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const functions = injector.get(IFunctionService)
  const executors = [
    new DvhFontFunction(store),
    new DvhColorFunction('fill', store),
    new DvhColorFunction('font', store),
    new DvhTableFunction(store),
  ]
  functions.registerExecutors(...executors)
  const infos = descriptions.map((d): IFunctionInfo => ({
    ...d,
    functionType: FunctionType.User,
    description: `DVH Tool · ${d.abstract}`,
  }))
  const descriptionHandle = injector.get(IDescriptionService).registerDescriptions(infos)
  // DVH.Table and Table: every function also answers to its short name
  const aliases = registerDvhAliases(injector, executors, infos)
  return {
    dispose() {
      aliases.dispose()
      descriptionHandle.dispose()
      functions.unregisterExecutors(...executors.map((executor) => executor.name))
    },
  }
}

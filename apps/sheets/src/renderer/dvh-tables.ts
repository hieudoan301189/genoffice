/**
 * DVH collections and tables in Sheets (P2).
 *
 * - A collection is a range with a title row (`_dvh.c.<id>`): columns come
 *   from the titles, rows from the cells below, formatting from the title and
 *   first data cells (Keep Source Style). The range grows and shrinks with the
 *   rows typed directly under it.
 * - A DVH.Table renders a collection as values and styles in one range
 *   (`_dvh.t.<id>`). Refresh inserts or deletes rows so the range always
 *   matches the data, and asks before overwriting cells edited by hand.
 */
import {
  collectionDefinedName,
  contentHash,
  effectiveTableColumns,
  newDvhId,
  renderTable,
  tableDefinedName,
  type DvhCellStyle,
  type DvhCollection,
  type DvhColumn,
  type DvhTable,
  type FieldType,
  type RenderedTable,
  type Scalar,
} from '@genoffice/dvh-model'
import { Dimension, ICommandService, type ICellData, type IStyleData } from '@univerjs/core'
import { SetRangeValuesMutation, type ISetRangeValuesMutationParams } from '@univerjs/sheets'

import {
  dvhSheetsContext,
  dvhStateOf,
  installDvhName,
  quoteSheetName,
  recordDvhChange,
  type DvhSheetState,
} from './dvh-smart-data'
import { t } from './i18n/locale'
import { univerDefinedNames } from './univer-sync'
import type { UniverRuntime } from './univer-state'

/** A zero-based block of cells on one sheet. */
export interface SheetRect {
  readonly sheetName: string
  readonly row: number
  readonly column: number
  readonly rows: number
  readonly columns: number
}

function columnLetters(index: number): string {
  let n = index + 1
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

function columnIndex(letters: string): number {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

/** `Sheet1!$A$1:$C$4` */
export function rectFormula(rect: SheetRect): string {
  const start = `$${columnLetters(rect.column)}$${rect.row + 1}`
  const end = `$${columnLetters(rect.column + rect.columns - 1)}$${rect.row + rect.rows}`
  return `${quoteSheetName(rect.sheetName)}!${start}:${end}`
}

/** `Sheet1!$A$1:$C$4`, `'Dữ liệu'!A1` → rect, or null (including `#REF!`). */
export function parseRectFormula(formula: string): SheetRect | null {
  const m =
    /^=?(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(
      formula.trim(),
    )
  if (!m) return null
  const sheetName = (m[1] ?? m[2] ?? '').replace(/''/g, "'")
  const c1 = columnIndex(m[3]!)
  const r1 = Number(m[4]) - 1
  const c2 = m[5] ? columnIndex(m[5]) : c1
  const r2 = m[6] ? Number(m[6]) - 1 : r1
  return {
    sheetName,
    row: Math.min(r1, r2),
    column: Math.min(c1, c2),
    rows: Math.abs(r2 - r1) + 1,
    columns: Math.abs(c2 - c1) + 1,
  }
}

function definedRect(runtime: UniverRuntime, name: string): SheetRect | null {
  const entry = univerDefinedNames(runtime).find((n) => n.getName() === name)
  return entry ? parseRectFormula(entry.getFormulaOrRefString()) : null
}

function setDefinedRect(runtime: UniverRuntime, name: string, rect: SheetRect): void {
  const existing = univerDefinedNames(runtime).find((n) => n.getName() === name)
  if (existing) existing.setRef(rectFormula(rect))
  else installDvhName(runtime, name, rectFormula(rect))
}

function sheetOf(runtime: UniverRuntime, sheetName: string) {
  const sheet = runtime.univerAPI.getActiveWorkbook()?.getSheetByName(sheetName)
  if (!sheet) throw new Error(`sheet ${sheetName} not found`)
  return sheet
}

const toScalar = (value: unknown): Scalar =>
  value === undefined || value === ''
    ? null
    : typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string'
      ? value
      : value === null
        ? null
        : String(value)

function readValues(runtime: UniverRuntime, rect: SheetRect): Scalar[][] {
  const values = sheetOf(runtime, rect.sheetName)
    .getRange(rect.row, rect.column, rect.rows, rect.columns)
    .getValues() as unknown[][]
  return values.map((row) => row.map(toScalar))
}

// ---------------- styles ----------------

const ALIGN_TO_UNIVER = { left: 1, center: 2, right: 3 } as const

/** The part of a Univer cell style a DVH table carries. */
export function dvhStyleFromUniver(style: IStyleData | null | undefined): DvhCellStyle {
  if (!style) return {}
  const out: DvhCellStyle = {}
  const hex = (rgb: unknown) =>
    typeof rgb === 'string' && /^#[0-9a-f]{6}$/i.test(rgb) ? rgb.toUpperCase() : undefined
  if (style.bl === 1) out.bold = true
  if (style.it === 1) out.italic = true
  const color = hex(style.cl?.rgb)
  if (color && color !== '#000000') out.color = color
  const fill = hex(style.bg?.rgb)
  if (fill && fill !== '#FFFFFF') out.fill = fill
  if (style.ht === 1) out.align = 'left'
  else if (style.ht === 2) out.align = 'center'
  else if (style.ht === 3) out.align = 'right'
  if (typeof style.fs === 'number') out.fontSize = style.fs
  const side = style.bd?.b ?? style.bd?.t ?? style.bd?.l ?? style.bd?.r
  if (side && side.s) {
    out.border = {
      color: hex(side.cl?.rgb) ?? '#000000',
      width: side.s >= 8 ? 'medium' : 'thin',
    }
  }
  return out
}

/** A DVH cell style (and number format) as a Univer cell style. */
export function univerStyleFromDvh(style: DvhCellStyle, numFmt?: string): IStyleData {
  const out: IStyleData = {}
  if (style.bold) out.bl = 1
  if (style.italic) out.it = 1
  if (style.color) out.cl = { rgb: style.color }
  if (style.fill) out.bg = { rgb: style.fill }
  if (style.align) out.ht = ALIGN_TO_UNIVER[style.align]
  if (style.fontSize) out.fs = style.fontSize
  if (style.border) {
    const side = { s: style.border.width === 'medium' ? 8 : 1, cl: { rgb: style.border.color } }
    out.bd = { t: side, b: side, l: side, r: side }
  }
  if (numFmt) out.n = { pattern: numFmt }
  return out
}

function cellStyle(runtime: UniverRuntime, sheetName: string, row: number, column: number) {
  return sheetOf(runtime, sheetName).getRange(row, column).getCellStyleData() as IStyleData | null
}

// ---------------- collections ----------------

function inferType(values: readonly Scalar[]): FieldType {
  const first = values.find((v) => v !== null)
  if (typeof first === 'number') return 'number'
  if (typeof first === 'boolean') return 'boolean'
  return 'text'
}

function columnKey(title: string, taken: Set<string>): string {
  const base =
    title
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/đ/gi, 'd')
      .replace(/[^A-Za-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '')
      .toLowerCase() || 'col'
  let key = base
  for (let n = 2; taken.has(key); n++) key = `${base}_${n}`
  taken.add(key)
  return key
}

/** Columns (titles, types, formats, widths) read from a range whose first row holds the titles. */
function readColumns(
  runtime: UniverRuntime,
  rect: SheetRect,
  values: Scalar[][],
  previous?: readonly DvhColumn[],
): DvhColumn[] {
  const sheet = sheetOf(runtime, rect.sheetName)
  const taken = new Set<string>(previous?.map((c) => c.key))
  const titles = Array.from({ length: rect.columns }, (_, i) =>
    values[0]?.[i] === null || values[0]?.[i] === undefined ? '' : String(values[0][i]),
  )
  const used = new Set<string>()
  /** a column keeps its id when its title is still there, or when it was renamed in place */
  const previousFor = (i: number): DvhColumn | undefined => {
    const byTitle = previous?.find((c) => c.title === titles[i] && !used.has(c.id))
    const atPosition = previous?.[i]
    const match =
      byTitle ??
      (atPosition && !used.has(atPosition.id) && !titles.includes(atPosition.title)
        ? atPosition
        : undefined)
    if (match) used.add(match.id)
    return match
  }
  return Array.from({ length: rect.columns }, (_, i) => {
    const title = titles[i]!
    const body = values.slice(1).map((row) => row[i] ?? null)
    const header = cellStyle(runtime, rect.sheetName, rect.row, rect.column + i)
    const first =
      rect.rows > 1 ? cellStyle(runtime, rect.sheetName, rect.row + 1, rect.column + i) : null
    const numFmt = first?.n?.pattern
    const old = previousFor(i)
    const column: DvhColumn = {
      id: old?.id ?? newDvhId('c').replace(/^c_/, 'col_'),
      key: old?.key ?? columnKey(title || `col ${i + 1}`, taken),
      title,
      type: inferType(body),
      ...(numFmt && numFmt !== 'General' ? { numFmt } : {}),
      headerStyle: dvhStyleFromUniver(header),
      style: dvhStyleFromUniver(first),
      width: sheet.getColumnWidth(rect.column + i),
    }
    return column
  })
}

/** Makes a range with a title row into a collection bound by `_dvh.c.<id>`. */
export function createCollection(
  runtime: UniverRuntime,
  dvh: DvhSheetState,
  options: { name: string; rect: SheetRect; source?: 'ui' | 'ai' },
): DvhCollection {
  const name = options.name.trim()
  if (!name) throw new Error('empty collection name')
  if (dvh.model?.collections.some((c) => c.name === name))
    throw new Error('duplicate collection name')
  if (options.rect.rows < 1) throw new Error('a collection needs a title row')
  const values = readValues(runtime, options.rect)
  const collection: DvhCollection = {
    id: newDvhId('c'),
    name,
    columns: readColumns(runtime, options.rect, values),
    rows: values.slice(1),
  }
  ensureModel(dvh).collections.push(collection)
  setDefinedRect(runtime, collectionDefinedName(collection.id), options.rect)
  recordDvhChange(
    dvh,
    'Data.CreateCollection',
    [
      {
        objectId: collection.id,
        path: '',
        before: null,
        after: {
          name,
          columns: collection.columns.map((c) => c.title),
          range: rectFormula(options.rect),
        },
      },
    ],
    options.source,
  )
  return collection
}

function ensureModel(dvh: DvhSheetState) {
  dvh.model ??= {
    docId: newDvhId('doc'),
    schemaVersion: 1,
    fields: [],
    collections: [],
    tables: [],
    links: [],
  }
  return dvh.model
}

const rowIsEmpty = (row: readonly Scalar[]) => row.every((v) => v === null)

/**
 * Re-reads a collection from its range. Rows typed directly under the range
 * join it, trailing empty rows leave it (Excel table auto-expand). Returns
 * false when the range is gone (its cells were deleted).
 */
/**
 * A collection as its range holds it right now: rows typed directly under the
 * range join it, a title typed right of the title row adds a column, trailing
 * empty rows and columns leave it (Excel table auto-expand). Reads only; null
 * when the range is gone (its cells were deleted).
 */
export function readCollectionState(
  runtime: UniverRuntime,
  collection: DvhCollection,
): { found: SheetRect; rect: SheetRect; columns: DvhColumn[]; rows: Scalar[][] } | null {
  const found = definedRect(runtime, collectionDefinedName(collection.id))
  if (!found) return null
  let rect = found
  const sheet = sheetOf(runtime, rect.sheetName)
  const maxRows = sheet.getMaxRows()
  const rowAt = (r: number) => readValues(runtime, { ...rect, row: r, rows: 1 })[0] ?? []
  const maxColumns = sheet.getMaxColumns()
  const titleAt = (c: number) =>
    readValues(runtime, { ...rect, column: c, columns: 1, rows: 1 })[0]?.[0]
  while (rect.column + rect.columns < maxColumns && titleAt(rect.column + rect.columns) != null) {
    rect = { ...rect, columns: rect.columns + 1 }
  }
  while (rect.row + rect.rows < maxRows && !rowIsEmpty(rowAt(rect.row + rect.rows))) {
    rect = { ...rect, rows: rect.rows + 1 }
  }
  let values = readValues(runtime, rect)
  while (rect.rows > 1 && rowIsEmpty(values[values.length - 1]!)) {
    rect = { ...rect, rows: rect.rows - 1 }
    values = values.slice(0, -1)
  }
  while (rect.columns > 1 && values.every((row) => row[rect.columns - 1] === null)) {
    rect = { ...rect, columns: rect.columns - 1 }
    values = values.map((row) => row.slice(0, -1))
  }
  return {
    found,
    rect,
    columns: readColumns(runtime, rect, values, collection.columns),
    rows: values.slice(1),
  }
}

const columnSignature = (columns: readonly DvhColumn[]) =>
  contentHash(columns.map((c) => [c.title, c.style, c.headerStyle, c.numFmt ?? null]))

/** Re-reads a collection from its range into the model, records the change, follows the range. */
export function syncCollection(
  runtime: UniverRuntime,
  dvh: DvhSheetState,
  collection: DvhCollection,
): boolean {
  const state = readCollectionState(runtime, collection)
  if (!state) return false
  const { found, rect, columns, rows } = state
  if (rectFormula(rect) !== rectFormula(found)) {
    setDefinedRect(runtime, collectionDefinedName(collection.id), rect)
  }
  const changed =
    contentHash(rows) !== contentHash(collection.rows) ||
    columnSignature(columns) !== columnSignature(collection.columns)
  if (changed) {
    recordDvhChange(dvh, 'Spreadsheet.CollectionEdited', [
      {
        objectId: collection.id,
        path: 'rows',
        before: collection.rows.length,
        after: rows.length,
      },
    ])
  }
  collection.columns = columns
  collection.rows = rows
  return true
}

export function syncCollections(runtime: UniverRuntime, dvh: DvhSheetState): void {
  for (const collection of dvh.model?.collections ?? []) syncCollection(runtime, dvh, collection)
}

/** Where each collection lives now (null when its range was deleted). */
export function collectionRanges(runtime: UniverRuntime | null, dvh: DvhSheetState) {
  return (dvh.model?.collections ?? []).map((collection) => {
    const rect = runtime ? definedRect(runtime, collectionDefinedName(collection.id)) : null
    return { collection, rect, formula: rect ? rectFormula(rect) : null }
  })
}

// ---------------- tables ----------------

/** A new table definition showing every column of a collection. */
export function createTable(
  dvh: DvhSheetState,
  options: {
    name: string
    collectionId: string
    mode?: 'source' | 'destination' | undefined
    source?: 'ui' | 'ai' | undefined
  },
): DvhTable {
  const model = ensureModel(dvh)
  const collection = model.collections.find(
    (c) => c.id === options.collectionId || c.name === options.collectionId,
  )
  if (!collection) throw new Error(`unknown collection ${options.collectionId}`)
  const name = options.name.trim() || collection.name
  if (model.tables.some((tb) => tb.name === name)) throw new Error('duplicate table name')
  const table: DvhTable = {
    id: newDvhId('t'),
    name,
    source: { collectionId: collection.id },
    columns: collection.columns.map((c) => ({
      id: newDvhId('c').replace(/^c_/, 'tc_'),
      columnId: c.id,
      title: c.title,
    })),
    autoColumns: true,
    style: { mode: options.mode ?? 'source' },
    layout: { repeatHeader: true, keepRowsTogether: true, widths: 'page' },
  }
  model.tables.push(table)
  recordDvhChange(
    dvh,
    'Table.Create',
    [{ objectId: table.id, path: '', before: null, after: { name, collectionId: collection.id } }],
    options.source,
  )
  return table
}

/** set while DVH writes a table, so the managed-range warning ignores its own edits */
let rendering = false

export interface RenderResult {
  readonly rect: SheetRect
  readonly rows: number
  /** the range held hand edits and `force` was not set: nothing was written */
  readonly needsConfirm?: boolean
}

function cellData(grid: RenderedTable): ICellData[][] {
  return grid.rows.map((row) => {
    const out: ICellData[] = []
    for (const cell of row) {
      const s = univerStyleFromDvh(cell.style, cell.numFmt)
      out.push({ v: cell.value ?? '', s })
      // cells under a span repeat the style so the merged block keeps its borders
      for (let k = 1; k < cell.colSpan; k++) out.push({ v: '', s })
    }
    return out
  })
}

/**
 * Renders (first time: at `anchor`) or refreshes a table. Row count changes
 * become row inserts/deletes inside the table range, so content below moves
 * with it.
 */
export function renderTableInSheet(
  runtime: UniverRuntime,
  dvh: DvhSheetState,
  tableRef: string,
  options: {
    anchor?: { sheetName: string; row: number; column: number }
    force?: boolean | undefined
    source?: 'ui' | 'ai' | undefined
  } = {},
): RenderResult {
  const model = ensureModel(dvh)
  const table = model.tables.find((tb) => tb.id === tableRef || tb.name === tableRef)
  if (!table) throw new Error(`unknown table ${tableRef}`)
  if ('collectionId' in table.source) {
    const collection = model.collections.find(
      (c) => c.id === (table.source as { collectionId: string }).collectionId,
    )
    if (collection) syncCollection(runtime, dvh, collection)
  }
  const grid = renderTable(table, model.collections)
  const name = tableDefinedName(table.id)
  const existing = definedRect(runtime, name)
  const at = existing ?? (options.anchor ? { ...options.anchor, rows: 0, columns: 0 } : null)
  if (!at) throw new Error(`table ${table.name} has no range yet; give an anchor cell`)

  if (existing && table.lastRender && !options.force) {
    if (contentHash(readValues(runtime, existing)) !== table.lastRender.hash) {
      return { rect: existing, rows: existing.rows, needsConfirm: true }
    }
  }

  const sheet = sheetOf(runtime, at.sheetName)
  const height = grid.rows.length
  rendering = true
  try {
    if (existing) {
      // a source column added or removed: shift cells in the table's rows only
      if (grid.columnCount > existing.columns) {
        sheet
          .getRange(
            existing.row,
            existing.column + existing.columns,
            existing.rows,
            grid.columnCount - existing.columns,
          )
          .insertCells(Dimension.COLUMNS)
      } else if (grid.columnCount < existing.columns) {
        sheet
          .getRange(
            existing.row,
            existing.column + grid.columnCount,
            existing.rows,
            existing.columns - grid.columnCount,
          )
          .deleteCells(Dimension.COLUMNS)
      }
      // shift cells in the table's columns only (Excel's "insert table rows"):
      // whole-row inserts would also split a source range sitting beside the table
      const width = grid.columnCount
      if (height > existing.rows) {
        sheet
          .getRange(
            existing.row + existing.rows - 1,
            existing.column,
            height - existing.rows,
            width,
          )
          .insertCells(Dimension.ROWS)
      } else if (height < existing.rows) {
        sheet
          .getRange(existing.row + height, existing.column, existing.rows - height, width)
          .deleteCells(Dimension.ROWS)
      }
      try {
        sheet.getRange(existing.row, existing.column, height, grid.columnCount).breakApart()
      } catch {
        // nothing merged
      }
    } else if (sheet.getMaxRows() < at.row + height) {
      sheet.insertRowsAfter(sheet.getMaxRows() - 1, at.row + height - sheet.getMaxRows())
    }
    const rect: SheetRect = {
      sheetName: at.sheetName,
      row: at.row,
      column: at.column,
      rows: height,
      columns: grid.columnCount,
    }
    sheet.getRange(rect.row, rect.column, rect.rows, rect.columns).setValues(cellData(grid))
    grid.rows.forEach((row, r) => {
      let column = 0
      for (const cell of row) {
        if (cell.colSpan > 1)
          sheet.getRange(rect.row + r, rect.column + column, 1, cell.colSpan).merge()
        column += cell.colSpan
      }
    })
    // source widths on the first render, and for columns that just appeared
    if (table.style.mode === 'source' && 'collectionId' in table.source) {
      const collection = model.collections.find(
        (c) => c.id === (table.source as { collectionId: string }).collectionId,
      )
      effectiveTableColumns(table, collection).forEach((col, i) => {
        if (existing && i < existing.columns) return
        const px = collection?.columns.find((c) => c.id === col.columnId)?.width
        if (px) sheet.setColumnWidth(rect.column + i, px)
      })
    }
    setDefinedRect(runtime, name, rect)
    table.lastRender = {
      hash: contentHash(readValues(runtime, rect)),
      at: new Date().toISOString(),
    }
    recordDvhChange(
      dvh,
      existing ? 'Table.Refresh' : 'Table.Render',
      [
        {
          objectId: table.id,
          path: 'render',
          before: existing ? rectFormula(existing) : null,
          after: { range: rectFormula(rect), bodyRows: grid.bodyRowCount },
        },
      ],
      options.source,
    )
    return { rect, rows: grid.bodyRowCount }
  } finally {
    rendering = false
  }
}

/** Where each table is rendered now (null before its first render). */
export function tableRanges(runtime: UniverRuntime | null, dvh: DvhSheetState) {
  return (dvh.model?.tables ?? []).map((table) => {
    const rect = runtime ? definedRect(runtime, tableDefinedName(table.id)) : null
    return { table, rect, formula: rect ? rectFormula(rect) : null }
  })
}

/** Changes a table's style mode or destination styles (render again to show it). */
export function setTableStyle(
  dvh: DvhSheetState,
  tableRef: string,
  style: { [K in keyof DvhTable['style']]?: DvhTable['style'][K] | undefined },
  source: 'ui' | 'ai' = 'ui',
): DvhTable {
  const table = dvh.model?.tables.find((tb) => tb.id === tableRef || tb.name === tableRef)
  if (!table) throw new Error(`unknown table ${tableRef}`)
  const before = table.style
  const defined = Object.fromEntries(Object.entries(style).filter(([, v]) => v !== undefined))
  table.style = { ...table.style, ...defined }
  recordDvhChange(
    dvh,
    'Table.SetStyle',
    [{ objectId: table.id, path: 'style', before, after: table.style }],
    source,
  )
  return table
}

/** Warns once per table when someone types inside a range a DVH table manages. */
export function installDvhTableWatch(runtime: UniverRuntime): { dispose(): void } {
  const commands = runtime.univer.__getInjector().get(ICommandService)
  const warned = new Set<string>()
  const listener = commands.onCommandExecuted((command) => {
    if (rendering || command.id !== SetRangeValuesMutation.id) return
    const params = command.params as ISetRangeValuesMutationParams | undefined
    const ctx = dvhSheetsContext()
    const dvh = dvhStateOf(ctx?.getState())
    if (!params?.cellValue || !ctx || !dvh?.model?.tables.length) return
    const sheetName = runtime.univerAPI
      .getActiveWorkbook()
      ?.getSheetBySheetId(params.subUnitId)
      ?.getSheetName()
    if (!sheetName) return
    const cells: [number, number][] = []
    for (const [r, cols] of Object.entries(params.cellValue)) {
      for (const c of Object.keys(cols ?? {})) cells.push([Number(r), Number(c)])
    }
    for (const { table, rect } of tableRanges(runtime, dvh)) {
      if (!rect || rect.sheetName !== sheetName || warned.has(table.id)) continue
      const inside = cells.some(
        ([r, c]) =>
          r >= rect.row &&
          r < rect.row + rect.rows &&
          c >= rect.column &&
          c < rect.column + rect.columns,
      )
      if (!inside) continue
      warned.add(table.id)
      ctx.notify(t('dvhTableManagedToast', { name: table.name }))
    }
  })
  return { dispose: () => listener.dispose() }
}

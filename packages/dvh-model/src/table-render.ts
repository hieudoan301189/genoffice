/**
 * Renders a DVH.Table into a grid of styled cells (P2). Sheets writes the grid
 * as cell values and styles, Docs as a docx table: both read the same result,
 * so a table looks the same wherever it is rendered.
 */

import type { DvhCellStyle, DvhCollection, DvhColumn, DvhTable, Scalar, TableTotal } from './types'

export type RenderedCellKind = 'group' | 'header' | 'body' | 'total'

export interface RenderedCell {
  readonly value: Scalar
  readonly kind: RenderedCellKind
  readonly style: DvhCellStyle
  readonly numFmt?: string
  /** grid columns this cell covers (group header cells span several) */
  readonly colSpan: number
}

export interface RenderedTable {
  readonly tableId: string
  readonly rows: readonly (readonly RenderedCell[])[]
  readonly columnCount: number
  /** group rows plus the title row: the rows that repeat on every page */
  readonly headerRowCount: number
  readonly bodyRowCount: number
  /** relative column widths, summing to 1 */
  readonly widths: readonly number[]
}

/** Destination style defaults. Document data, so plain colors (CLAUDE.md rule 4). */
export const DEFAULT_TABLE_STYLE: {
  readonly header: DvhCellStyle
  readonly body: DvhCellStyle
  readonly total: DvhCellStyle
} = {
  header: {
    bold: true,
    fill: '#D9E2F3',
    align: 'center',
    border: { color: '#7F7F7F', width: 'thin' },
  },
  body: { border: { color: '#7F7F7F', width: 'thin' } },
  total: { bold: true, border: { color: '#7F7F7F', width: 'thin' } },
}

/** Excel's default column width in pixels, the unit of DvhColumn.width */
const DEFAULT_COLUMN_PX = 64

function total(kind: TableTotal, values: readonly Scalar[]): Scalar {
  if (kind === 'count') return values.filter((v) => v !== null && v !== '').length
  const numbers = values.filter((v): v is number => typeof v === 'number')
  if (numbers.length === 0) return null
  if (kind === 'sum') return numbers.reduce((a, b) => a + b, 0)
  if (kind === 'avg') return numbers.reduce((a, b) => a + b, 0) / numbers.length
  if (kind === 'min') return Math.min(...numbers)
  if (kind === 'max') return Math.max(...numbers)
  return null
}

/** The columns a table shows now: with autoColumns, every collection column with its current title. */
export function effectiveTableColumns(
  table: DvhTable,
  collection: DvhCollection | undefined,
): DvhTable['columns'] {
  if (!table.autoColumns || !collection) return table.columns
  return collection.columns.map((c) => {
    const own = table.columns.find((col) => col.columnId === c.id)
    return { ...(own ?? { id: `tc_${c.id}`, columnId: c.id }), title: c.title }
  })
}

export function renderTable(input: DvhTable, collections: readonly DvhCollection[]): RenderedTable {
  const collection =
    'collectionId' in input.source
      ? collections.find((c) => c.id === (input.source as { collectionId: string }).collectionId)
      : undefined
  if ('collectionId' in input.source && !collection) {
    throw new Error(`table ${input.name}: collection ${input.source.collectionId} not found`)
  }
  const table: DvhTable = { ...input, columns: effectiveTableColumns(input, collection) }
  const dataRows = collection ? collection.rows : (table.source as { rows: Scalar[][] }).rows
  const sourceIndex = table.columns.map((col, i) =>
    collection ? collection.columns.findIndex((c) => c.id === col.columnId) : i,
  )
  const sourceColumn = (i: number): DvhColumn | undefined =>
    collection && sourceIndex[i]! >= 0 ? collection.columns[sourceIndex[i]!] : undefined
  const source = table.style.mode === 'source'

  const headerStyle = (i: number): DvhCellStyle =>
    source
      ? (sourceColumn(i)?.headerStyle ?? {})
      : { ...DEFAULT_TABLE_STYLE.header, ...table.style.header }
  const bodyStyle = (i: number, value: Scalar, band: boolean): DvhCellStyle => {
    const col = table.columns[i]!
    const base = source
      ? (sourceColumn(i)?.style ?? {})
      : { ...DEFAULT_TABLE_STYLE.body, ...table.style.body }
    const align =
      col.align ?? base.align ?? (!source && typeof value === 'number' ? 'right' : undefined)
    return {
      ...base,
      ...(align ? { align } : {}),
      ...(band && table.style.bandFill ? { fill: table.style.bandFill } : {}),
    }
  }
  const numFmtOf = (i: number) => table.columns[i]!.numFmt ?? sourceColumn(i)?.numFmt

  const rows: RenderedCell[][] = []
  for (const group of table.headerGroups ?? []) {
    const covered = group.reduce((sum, cell) => sum + cell.span, 0)
    if (covered !== table.columns.length) {
      // a source column was added or removed under a grouped header: show plain titles
      if (table.autoColumns) break
      throw new Error(
        `table ${table.name}: a header group spans ${covered} of ${table.columns.length} columns`,
      )
    }
    let column = 0
    rows.push(
      group.map((cell) => {
        const style = headerStyle(column)
        column += cell.span
        return { value: cell.title, kind: 'group', style, colSpan: cell.span }
      }),
    )
  }
  rows.push(
    table.columns.map((col, i) => ({
      value: col.title,
      kind: 'header',
      style: headerStyle(i),
      colSpan: 1,
    })),
  )
  const headerRowCount = rows.length

  const bodyValues = dataRows.map((row) =>
    table.columns.map((_, i) => (sourceIndex[i]! >= 0 ? (row[sourceIndex[i]!] ?? null) : null)),
  )
  bodyValues.forEach((values, r) => {
    rows.push(
      values.map((value, i) => {
        const numFmt = numFmtOf(i)
        return {
          value,
          kind: 'body',
          style: bodyStyle(i, value, r % 2 === 1),
          ...(numFmt ? { numFmt } : {}),
          colSpan: 1,
        }
      }),
    )
  })

  if (table.totalRow) {
    const labelAt = Math.max(
      0,
      table.columns.findIndex((c) => !c.total || c.total === 'none'),
    )
    rows.push(
      table.columns.map((col, i) => {
        const kind = col.total ?? 'none'
        const value =
          kind !== 'none'
            ? total(
                kind,
                bodyValues.map((v) => v[i] ?? null),
              )
            : i === labelAt
              ? table.totalRow!.label
              : null
        const numFmt = kind === 'count' ? undefined : numFmtOf(i)
        const base = source
          ? { ...(sourceColumn(i)?.style ?? {}), bold: true }
          : { ...DEFAULT_TABLE_STYLE.total, ...table.style.total }
        const align = kind !== 'none' ? (col.align ?? base.align ?? 'right') : base.align
        return {
          value,
          kind: 'total',
          style: { ...base, ...(align ? { align } : {}) },
          ...(numFmt ? { numFmt } : {}),
          colSpan: 1,
        }
      }),
    )
  }

  const weights = table.columns.map((col, i) => {
    const px = sourceColumn(i)?.width
    return col.width ?? (px ? px / DEFAULT_COLUMN_PX : 1)
  })
  const weightSum = weights.reduce((a, b) => a + b, 0)
  return {
    tableId: table.id,
    rows,
    columnCount: table.columns.length,
    headerRowCount,
    bodyRowCount: bodyValues.length,
    widths: weights.map((w) => w / weightSum),
  }
}

// ---- number formats ----

interface FormatSection {
  prefix: string
  suffix: string
  /** the digit pattern, e.g. `#,##0.00` */
  digits: string
  percent: boolean
  date: boolean
  raw: string
}

/** Splits the first section of an Excel number format into literal text around one digit pattern. */
function parseSection(numFmt: string): FormatSection {
  const section = numFmt.split(';')[0] ?? ''
  let prefix = ''
  let suffix = ''
  let digits = ''
  let percent = false
  for (let i = 0; i < section.length; i++) {
    const ch = section[i]!
    let literal: string | null = null
    if (ch === '"') {
      const end = section.indexOf('"', i + 1)
      literal = section.slice(i + 1, end < 0 ? undefined : end)
      i = end < 0 ? section.length : end
    } else if (ch === '\\') {
      literal = section[++i] ?? ''
    } else if (ch === '_' || ch === '*') {
      i++
      literal = ch === '_' ? ' ' : ''
    } else if (ch === '[') {
      i = Math.max(i, section.indexOf(']', i))
      literal = ''
    } else if ('0#?,.'.includes(ch) && (digits === '' ? suffix === '' : suffix === '')) {
      digits += ch
    } else if (ch === '%') {
      percent = true
      literal = '%'
    } else {
      literal = ch
    }
    if (literal !== null) {
      if (digits === '') prefix += literal
      else suffix += literal
    }
  }
  const date = digits === '' && /[dmyhs]/i.test(section.replace(/"[^"]*"/g, ''))
  return { prefix, suffix, digits, percent, date, raw: section }
}

const pad = (n: number, width: number) => String(n).padStart(width, '0')

/** Excel serial (1900 system) to a date string following d/m/y tokens of the format. */
function formatDate(serial: number, pattern: string): string {
  const date = new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000))
  const parts: Record<string, string> = {
    yyyy: String(date.getUTCFullYear()),
    yy: pad(date.getUTCFullYear() % 100, 2),
    mm: pad(date.getUTCMonth() + 1, 2),
    m: String(date.getUTCMonth() + 1),
    dd: pad(date.getUTCDate(), 2),
    d: String(date.getUTCDate()),
  }
  return pattern
    .replace(/"([^"]*)"/g, '$1')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/yyyy|yy|mm|m|dd|d/gi, (token) => parts[token.toLowerCase()] ?? token)
}

/**
 * The text a cell shows for a value and an Excel number format, for renderers
 * that have no spreadsheet engine (Docs). Covers digit patterns with grouping
 * and decimals, percent, quoted literal text and d/m/y dates.
 */
export function formatCellValue(value: Scalar, numFmt?: string, locale = 'en-US'): string {
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'string') return value
  if (!numFmt || /^general$/i.test(numFmt.trim())) {
    return new Intl.NumberFormat(locale, { maximumFractionDigits: 10, useGrouping: false }).format(
      value,
    )
  }
  const f = parseSection(numFmt)
  if (f.date) return formatDate(value, f.raw)
  if (f.digits === '') return f.prefix + f.suffix
  const [intPart, fracPart = ''] = f.digits.split('.')
  const minFraction = (fracPart.match(/0/g) ?? []).length
  const maxFraction = (fracPart.match(/[0#?]/g) ?? []).length
  const text = new Intl.NumberFormat(locale, {
    minimumFractionDigits: minFraction,
    maximumFractionDigits: maxFraction,
    useGrouping: (intPart ?? '').includes(','),
  }).format(f.percent ? value * 100 : value)
  return f.prefix + text + f.suffix
}

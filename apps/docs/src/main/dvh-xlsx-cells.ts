/**
 * Reads DVH Smart Data straight from a workbook's cells (P3 automatic links).
 *
 * The model part holds the values of the last DVH save; a workbook edited and
 * saved in Microsoft Excel keeps that part untouched while its cells change.
 * The hidden names `_dvh.f.<id>` (one cell) and `_dvh.c.<id>` (title row +
 * data rows) still point at the data, so a linked document refreshes from the
 * cells: values, number formats, the title/first-row formatting and widths.
 */
import JSZip from 'jszip'
import {
  COLLECTION_NAME_PREFIX,
  FIELD_NAME_PREFIX,
  fieldValueFromText,
  newDvhId,
  type DvhCellStyle,
  type DvhCollection,
  type DvhColumn,
  type DvhModel,
  type FieldType,
  type Scalar,
} from '@genoffice/dvh-model'

interface Rect {
  sheetName: string
  row: number
  column: number
  rows: number
  columns: number
}

interface CellRecord {
  value: Scalar
  styleIndex: number
}

const decode = (text: string) =>
  text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n: string) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n: string) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, '&')

const attr = (tag: string, name: string) => {
  const m = new RegExp(`\\s${name}="([^"]*)"`).exec(tag)
  return m ? decode(m[1]!) : undefined
}

export function columnIndex(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

/** `Sheet1!$A$1:$C$4` → rect (zero-based), or null. */
export function parseRef(ref: string): Rect | null {
  const m =
    /^=?(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/.exec(
      ref.trim(),
    )
  if (!m) return null
  const c1 = columnIndex(m[3]!)
  const r1 = Number(m[4]) - 1
  const c2 = m[5] ? columnIndex(m[5]) : c1
  const r2 = m[6] ? Number(m[6]) - 1 : r1
  return {
    sheetName: (m[1] ?? m[2] ?? '').replace(/''/g, "'"),
    row: Math.min(r1, r2),
    column: Math.min(c1, c2),
    rows: Math.abs(r2 - r1) + 1,
    columns: Math.abs(c2 - c1) + 1,
  }
}

// ---------------- styles ----------------

/** Built-in number formats Excel does not write into styles.xml. */
const BUILTIN_NUMFMTS: Record<number, string> = {
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  14: 'dd/mm/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  22: 'dd/mm/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  49: '@',
}

interface Styles {
  numFmt(styleIndex: number): string | undefined
  cellStyle(styleIndex: number): DvhCellStyle
}

function section(xml: string, name: string): string {
  const m = new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)</${name}>`).exec(xml)
  return m ? m[1]! : ''
}
function items(xml: string, name: string): string[] {
  return [...xml.matchAll(new RegExp(`<${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</${name}>)`, 'g'))].map(
    (m) => m[0],
  )
}

function themeColors(themeXml: string): string[] {
  const scheme = section(themeXml, 'a:clrScheme')
  const order = [
    'dk1',
    'lt1',
    'dk2',
    'lt2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
  ]
  const colors = order.map((slot) => {
    const block = section(scheme, `a:${slot}`)
    return (
      /<a:srgbClr\b[^>]*val="([0-9A-Fa-f]{6})"/.exec(block)?.[1] ??
      /<a:sysClr\b[^>]*lastClr="([0-9A-Fa-f]{6})"/.exec(block)?.[1] ??
      null
    )
  })
  // Excel's theme index swaps the first two pairs: 0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2
  return [colors[1], colors[0], colors[3], colors[2], ...colors.slice(4)].map((c) =>
    c ? `#${c.toUpperCase()}` : '',
  )
}

function readStyles(stylesXml: string, themeXml: string): Styles {
  const theme = themeColors(themeXml)
  const color = (tag: string | undefined): string | undefined => {
    if (!tag) return undefined
    const rgb = attr(tag, 'rgb')
    if (rgb && /^[0-9A-Fa-f]{8}$/.test(rgb)) return `#${rgb.slice(2).toUpperCase()}`
    if (rgb && /^[0-9A-Fa-f]{6}$/.test(rgb)) return `#${rgb.toUpperCase()}`
    const index = attr(tag, 'theme')
    if (index !== undefined && attr(tag, 'tint') === undefined)
      return theme[Number(index)] || undefined
    return undefined
  }
  const custom = new Map<number, string>()
  for (const fmt of items(section(stylesXml, 'numFmts'), 'numFmt')) {
    custom.set(Number(attr(fmt, 'numFmtId')), attr(fmt, 'formatCode') ?? '')
  }
  const fonts = items(section(stylesXml, 'fonts'), 'font').map((font) => {
    const out: DvhCellStyle = {}
    if (/<b(\s[^>]*)?\/>/.test(font) && !/<b\s+val="(0|false)"/.test(font)) out.bold = true
    if (/<i(\s[^>]*)?\/>/.test(font) && !/<i\s+val="(0|false)"/.test(font)) out.italic = true
    const c = color(/<color\b[^>]*\/?>/.exec(font)?.[0])
    if (c && c !== '#000000') out.color = c
    const size = /<sz\b[^>]*val="([\d.]+)"/.exec(font)?.[1]
    if (size) out.fontSize = Number(size)
    return out
  })
  const fills = items(section(stylesXml, 'fills'), 'fill').map((fill) => {
    if (!/patternType="solid"/.test(fill)) return undefined
    const c = color(/<fgColor\b[^>]*\/?>/.exec(fill)?.[0])
    return c && c !== '#FFFFFF' ? c : undefined
  })
  const borders = items(section(stylesXml, 'borders'), 'border').map((border) => {
    for (const side of ['bottom', 'top', 'left', 'right']) {
      const m = new RegExp(`<${side}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${side}>)`).exec(border)
      const style = m ? attr(`<x${m[1]}>`, 'style') : undefined
      if (!style || style === 'none') continue
      return {
        color: color(/<color\b[^>]*\/?>/.exec(m![2] ?? '')?.[0]) ?? '#000000',
        width: /medium|thick|double/.test(style) ? ('medium' as const) : ('thin' as const),
      }
    }
    return undefined
  })
  const xfs = items(section(stylesXml, 'cellXfs'), 'xf').map((xf) => {
    const head = /<xf\b[^>]*>/.exec(xf)?.[0] ?? xf
    const align = attr(/<alignment\b[^>]*\/?>/.exec(xf)?.[0] ?? '', 'horizontal')
    return {
      numFmtId: Number(attr(head, 'numFmtId') ?? 0),
      fontId: Number(attr(head, 'fontId') ?? 0),
      fillId: Number(attr(head, 'fillId') ?? 0),
      borderId: Number(attr(head, 'borderId') ?? 0),
      align:
        align === 'left' || align === 'center' || align === 'right'
          ? (align as 'left' | 'center' | 'right')
          : undefined,
    }
  })
  return {
    numFmt(styleIndex) {
      const id = xfs[styleIndex]?.numFmtId ?? 0
      if (id === 0) return undefined
      return custom.get(id) ?? BUILTIN_NUMFMTS[id]
    },
    cellStyle(styleIndex) {
      const xf = xfs[styleIndex]
      if (!xf) return {}
      const fill = fills[xf.fillId]
      const border = borders[xf.borderId]
      return {
        ...(fonts[xf.fontId] ?? {}),
        ...(fill ? { fill } : {}),
        ...(xf.align ? { align: xf.align } : {}),
        ...(border ? { border } : {}),
      }
    },
  }
}

// ---------------- sheets ----------------

function sharedStrings(xml: string): string[] {
  return items(xml, 'si').map((si) =>
    [...si.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decode(m[1]!)).join(''),
  )
}

/** The cells of a worksheet inside the given rows/columns. */
function readCells(
  sheetXml: string,
  strings: readonly string[],
  within: (row: number, column: number) => boolean,
): Map<string, CellRecord> {
  const cells = new Map<string, CellRecord>()
  for (const m of sheetXml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const ref = attr(`<c${m[1]}>`, 'r')
    const pos = ref ? /^([A-Z]{1,3})(\d+)$/.exec(ref) : null
    if (!pos) continue
    const row = Number(pos[2]) - 1
    const column = columnIndex(pos[1]!)
    if (!within(row, column)) continue
    const type = attr(`<c${m[1]}>`, 't')
    const body = m[2] ?? ''
    const raw = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1]
    let value: Scalar = null
    if (type === 's' && raw !== undefined) value = strings[Number(raw)] ?? null
    else if (type === 'inlineStr')
      value = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => decode(t[1]!)).join('')
    else if (type === 'b' && raw !== undefined) value = raw === '1'
    else if ((type === 'str' || type === 'e') && raw !== undefined) value = decode(raw)
    else if (raw !== undefined && raw !== '') value = Number(raw)
    cells.set(`${row}:${column}`, {
      value: value === '' ? null : value,
      styleIndex: Number(attr(`<c${m[1]}>`, 's') ?? 0),
    })
  }
  return cells
}

function columnWidths(sheetXml: string): Map<number, number> {
  const widths = new Map<number, number>()
  for (const col of items(section(sheetXml, 'cols'), 'col')) {
    const width = Number(attr(col, 'width'))
    if (!Number.isFinite(width)) continue
    for (let c = Number(attr(col, 'min')) - 1; c <= Number(attr(col, 'max')) - 1; c++) {
      widths.set(c, Math.round(width * 7 + 5))
    }
  }
  return widths
}

function inferType(values: readonly Scalar[]): FieldType {
  const first = values.find((v) => v !== null)
  if (typeof first === 'number') return 'number'
  if (typeof first === 'boolean') return 'boolean'
  return 'text'
}

/** Column ids survive when their title is still there, or when renamed in place. */
function matchColumns(previous: readonly DvhColumn[], titles: readonly string[]) {
  const used = new Set<string>()
  return titles.map((title, i) => {
    const byTitle = previous.find((c) => c.title === title && !used.has(c.id))
    const atPosition = previous[i]
    const match =
      byTitle ??
      (atPosition && !used.has(atPosition.id) && !titles.includes(atPosition.title)
        ? atPosition
        : undefined)
    if (match) used.add(match.id)
    return match
  })
}

export interface WorkbookLayout {
  /** every `_dvh.f.*` / `_dvh.c.*` name → its reference text */
  names: Map<string, string>
  /** relationship id → package path */
  targets: Map<string, string>
  /** sheet name → worksheet part path */
  sheetPaths: Map<string, string>
  workbookXml: string
}

/** The workbook's DVH names and where each sheet lives in the package. */
export async function workbookLayout(zip: JSZip): Promise<WorkbookLayout> {
  const text = async (path: string) => (await zip.file(path)?.async('string')) ?? ''
  const workbookXml = await text('xl/workbook.xml')
  const names = new Map<string, string>()
  for (const m of workbookXml.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)) {
    const name = attr(`<x${m[1]}>`, 'name')
    if (name?.startsWith(FIELD_NAME_PREFIX) || name?.startsWith(COLLECTION_NAME_PREFIX)) {
      names.set(name, decode(m[2]!))
    }
  }
  const rels = await text('xl/_rels/workbook.xml.rels')
  const targets = new Map<string, string>()
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], 'Id')
    const target = attr(m[0], 'Target')
    if (id && target) {
      targets.set(
        id,
        target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`,
      )
    }
  }
  const sheetPaths = new Map<string, string>()
  for (const m of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = attr(m[0], 'name')
    const rid = attr(m[0], 'r:id')
    const path = rid ? targets.get(rid) : undefined
    if (name && path) sheetPaths.set(name, path)
  }
  return { names, targets, sheetPaths, workbookXml }
}

/**
 * The model with its field values and collections re-read from the cells the
 * hidden DVH names point at. Objects whose names are missing keep the values
 * the model part holds.
 */
export async function modelFromCells(bytes: Uint8Array, model: DvhModel): Promise<DvhModel> {
  const zip = await JSZip.loadAsync(bytes)
  const text = async (path: string) => (await zip.file(path)?.async('string')) ?? ''
  const { names, targets, sheetPaths } = await workbookLayout(zip)
  if (names.size === 0) return model

  const strings = sharedStrings(await text('xl/sharedStrings.xml'))
  const themePath =
    [...targets.values()].find((p) => /theme\d*\.xml$/.test(p)) ?? 'xl/theme/theme1.xml'
  const styles = readStyles(await text('xl/styles.xml'), await text(themePath))

  // read each sheet once, keeping only the cells a DVH name covers
  const wanted = new Map<string, Rect[]>()
  for (const ref of names.values()) {
    const rect = parseRef(ref)
    if (rect) wanted.set(rect.sheetName, [...(wanted.get(rect.sheetName) ?? []), rect])
  }
  const sheets = new Map<string, { cells: Map<string, CellRecord>; widths: Map<number, number> }>()
  for (const [sheetName, rects] of wanted) {
    const path = sheetPaths.get(sheetName)
    if (!path) continue
    const xml = await text(path)
    // collections may have grown below / right of their range: Excel does not stretch the name
    const within = (row: number, column: number) =>
      rects.some((r) => row >= r.row && column >= r.column)
    sheets.set(sheetName, { cells: readCells(xml, strings, within), widths: columnWidths(xml) })
  }
  const cellAt = (rect: Rect, row: number, column: number) =>
    sheets.get(rect.sheetName)?.cells.get(`${row}:${column}`)

  const next: DvhModel = structuredClone(model)
  for (const field of next.fields) {
    const rect = parseRef(names.get(`${FIELD_NAME_PREFIX}${field.id}`) ?? '')
    if (!rect || !sheets.has(rect.sheetName)) continue
    const cell = cellAt(rect, rect.row, rect.column)
    const value = cell?.value ?? null
    field.value = typeof value === 'string' ? fieldValueFromText(field.type, value) : value
  }
  next.collections = next.collections.map((collection): DvhCollection => {
    const named = parseRef(names.get(`${COLLECTION_NAME_PREFIX}${collection.id}`) ?? '')
    const sheet = named ? sheets.get(named.sheetName) : undefined
    if (!named || !sheet) return collection
    // same rule as Sheets: titles typed right of the title row and rows typed under
    // the range join it; an empty last row or column leaves it
    const filled = (row: number, column: number) => cellAt(named, row, column)?.value != null
    const rect = { ...named }
    while (filled(rect.row, rect.column + rect.columns)) rect.columns++
    const rowFilled = (row: number) =>
      Array.from({ length: rect.columns }, (_, c) => filled(row, rect.column + c)).some(Boolean)
    const columnFilled = (column: number) =>
      Array.from({ length: rect.rows }, (_, r) => filled(rect.row + r, column)).some(Boolean)
    while (rowFilled(rect.row + rect.rows)) rect.rows++
    while (rect.rows > 1 && !rowFilled(rect.row + rect.rows - 1)) rect.rows--
    while (rect.columns > 1 && !columnFilled(rect.column + rect.columns - 1)) rect.columns--
    const grid = Array.from({ length: rect.rows }, (_, r) =>
      Array.from(
        { length: rect.columns },
        (_, c) => cellAt(rect, rect.row + r, rect.column + c)?.value ?? null,
      ),
    )
    const titles = (grid[0] ?? []).map((v) => (v === null ? '' : String(v)))
    const previous = matchColumns(collection.columns, titles)
    const columns = titles.map((title, i): DvhColumn => {
      const header = cellAt(rect, rect.row, rect.column + i)
      const first = rect.rows > 1 ? cellAt(rect, rect.row + 1, rect.column + i) : undefined
      const numFmt = first ? styles.numFmt(first.styleIndex) : undefined
      const width = sheet.widths.get(rect.column + i)
      return {
        id: previous[i]?.id ?? newDvhId('c').replace(/^c_/, 'col_'),
        key: previous[i]?.key ?? `col_${i + 1}`,
        title,
        type: inferType(grid.slice(1).map((row) => row[i] ?? null)),
        ...(numFmt && numFmt !== '@' ? { numFmt } : {}),
        headerStyle: header ? styles.cellStyle(header.styleIndex) : {},
        style: first ? styles.cellStyle(first.styleIndex) : {},
        ...(width ? { width } : previous[i]?.width ? { width: previous[i]!.width } : {}),
      }
    })
    return { ...collection, columns, rows: grid.slice(1) }
  })
  return next
}

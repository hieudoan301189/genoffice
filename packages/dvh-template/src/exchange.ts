/**
 * Data exchange with QLCL-DVH (P6, one way at a time; two-way sync is P10).
 *
 * Two carriers:
 * - a versioned JSON document (`dvh-exchange`, version 1) whose JSON Schema
 *   is exported from the dvh-model schemas, for tools that speak JSON;
 * - a workbook in QLCL's own layout, usable by QLCL-DVH unchanged: sheet
 *   `ThongTin` holds name/value pairs (→ fields), sheet `Data` a header row
 *   and records (→ collection `Data`); any other sheet with a header row
 *   becomes a collection of its name.
 */
import JSZip from 'jszip'
import { z } from 'zod'
import {
  dvhModelSchema,
  emptyModel,
  fieldValueFromText,
  newDvhId,
  type DvhCollection,
  type DvhModel,
  type FieldType,
  type Scalar,
} from '@genoffice/dvh-model'
import { keyOf } from './convert'
import { deterministicZip } from './fill'
import { escapeText } from './ooxml'

export const EXCHANGE_FORMAT = 'dvh-exchange'
export const EXCHANGE_VERSION = 1

export const exchangeDocumentSchema = z
  .object({
    format: z.literal(EXCHANGE_FORMAT),
    version: z.literal(EXCHANGE_VERSION),
    /** who wrote it (`DVH Office`, `QLCL-DVH`…) */
    producer: z.string().optional(),
    exportedAt: z.string(),
    model: dvhModelSchema,
  })
  .strict()
export type ExchangeDocument = z.infer<typeof exchangeDocumentSchema>

/** The JSON Schema of the exchange document, for QLCL-DVH and other producers. */
export function exchangeJsonSchema(): unknown {
  return z.toJSONSchema(exchangeDocumentSchema, { unrepresentable: 'any' })
}

export function exportExchangeJson(
  model: DvhModel,
  exportedAt: string,
  producer = 'DVH Office',
): string {
  const doc: ExchangeDocument = {
    format: EXCHANGE_FORMAT,
    version: EXCHANGE_VERSION,
    producer,
    exportedAt,
    model,
  }
  return JSON.stringify(doc, null, 2)
}

/** Reads an exchange document; throws a readable error on anything else. */
export function importExchangeJson(text: string): DvhModel {
  const parsed = exchangeDocumentSchema.safeParse(JSON.parse(text))
  if (!parsed.success)
    throw new Error(
      `not a ${EXCHANGE_FORMAT} v${EXCHANGE_VERSION} document: ${parsed.error.message}`,
    )
  return parsed.data.model
}

// ---------------- QLCL workbook ----------------

const decode = (s: string) =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n: string) => String.fromCodePoint(Number(n)))
    .replace(/&amp;/g, '&')

function columnIndex(letters: string): number {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

function columnLetters(index: number): string {
  let n = index + 1
  let out = ''
  while (n > 0) {
    out = String.fromCharCode(65 + ((n - 1) % 26)) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/** Every sheet of a workbook as a grid of values (formulas give their cached value). */
export async function readWorkbookGrids(bytes: Uint8Array): Promise<Map<string, Scalar[][]>> {
  const zip = await JSZip.loadAsync(bytes)
  const text = async (p: string) => (await zip.file(p)?.async('string')) ?? ''
  const workbook = await text('xl/workbook.xml')
  const rels = await text('xl/_rels/workbook.xml.rels')
  const targets = new Map<string, string>()
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[0])?.[1]
    const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1]
    if (id && target) targets.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`)
  }
  const strings = [...(await text('xl/sharedStrings.xml')).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(
    (si) => [...si[1]!.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => decode(t[1]!)).join(''),
  )
  const grids = new Map<string, Scalar[][]>()
  for (const m of workbook.matchAll(/<sheet\b[^>]*>/g)) {
    const name = decode(/\bname="([^"]*)"/.exec(m[0])?.[1] ?? '')
    const rid = /\br:id="([^"]+)"/.exec(m[0])?.[1]
    const path = rid ? targets.get(rid) : undefined
    if (!path) continue
    const grid: Scalar[][] = []
    for (const c of (await text(path)).matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = /\br="([A-Z]+)(\d+)"/.exec(c[1]!)
      if (!ref) continue
      const type = /\bt="([^"]+)"/.exec(c[1]!)?.[1]
      const body = c[2] ?? ''
      const raw = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1]
      let value: Scalar = null
      if (type === 's' && raw !== undefined) value = strings[Number(raw)] ?? null
      else if (type === 'inlineStr')
        value = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => decode(t[1]!)).join('')
      else if (type === 'b' && raw !== undefined) value = raw === '1'
      else if ((type === 'str' || type === 'e') && raw !== undefined) value = decode(raw)
      else if (raw !== undefined && raw !== '') value = Number(raw)
      const r = Number(ref[2]) - 1
      const col = columnIndex(ref[1]!)
      ;(grid[r] ??= [])[col] = value === '' ? null : value
    }
    grids.set(
      name,
      grid.map((row) => Array.from({ length: row?.length ?? 0 }, (_, i) => row?.[i] ?? null)),
    )
  }
  return grids
}

const inferType = (values: readonly Scalar[]): FieldType => {
  const first = values.find((v) => v !== null)
  return typeof first === 'number' ? 'number' : typeof first === 'boolean' ? 'boolean' : 'text'
}

/** Header row + records → a collection (empty trailing rows dropped). */
function collectionFrom(
  name: string,
  grid: readonly Scalar[][],
  previous?: DvhCollection,
): DvhCollection | null {
  const header = (grid[0] ?? []).map((v) => (v === null ? '' : String(v).trim()))
  while (header.length && header.at(-1) === '') header.pop()
  if (header.length === 0) return null
  const rows = grid
    .slice(1)
    .map((row) => header.map((_, i) => row[i] ?? null))
    .filter((row) => row.some((v) => v !== null))
  return {
    id: previous?.id ?? newDvhId('c'),
    name,
    columns: header.map((title, i) => {
      const old = previous?.columns.find((c) => c.title === title)
      return {
        id: old?.id ?? newDvhId('c').replace(/^c_/, 'col_'),
        key: old?.key ?? keyOf(title || `cot ${i + 1}`),
        title,
        type: inferType(rows.map((r) => r[i] ?? null)),
      }
    }),
    rows,
  }
}

/**
 * QLCL workbook → Smart Data. With `into`, objects of the same name keep
 * their ids (fields by name, collections by name, columns by title), so
 * documents linked to them stay linked across imports.
 */
export async function importQlclWorkbook(bytes: Uint8Array, into?: DvhModel): Promise<DvhModel> {
  const grids = await readWorkbookGrids(bytes)
  const model: DvhModel = structuredClone(into ?? emptyModel(newDvhId('doc')))
  const info = grids.get('ThongTin') ?? []
  // a header row ("Tên | Giá trị", "Name | Value") is not a field
  const headerRow = /^(giá trị|gia tri|value)$/i.test(String(info[0]?.[1] ?? '').trim())
  for (const row of headerRow ? info.slice(1) : info) {
    const name = row[0] === null || row[0] === undefined ? '' : String(row[0]).trim()
    if (!name) continue
    const value = row[1] ?? null
    const existing = model.fields.find((f) => f.name === name)
    if (existing) {
      existing.value = typeof value === 'string' ? fieldValueFromText(existing.type, value) : value
    } else {
      model.fields.push({
        id: newDvhId('f'),
        name,
        type: inferType([value]),
        value,
        access: 'readwrite',
      })
    }
  }
  for (const [name, grid] of grids) {
    if (name === 'ThongTin') continue
    const previous = model.collections.find((c) => c.name === name)
    const collection = collectionFrom(name, grid, previous)
    if (!collection) continue
    model.collections = [...model.collections.filter((c) => c.name !== name), collection]
  }
  return model
}

function sheetXml(rows: readonly (readonly Scalar[])[]): string {
  const body = rows
    .map((row, r) => {
      const cells = row
        .map((v, c) => {
          const ref = `${columnLetters(c)}${r + 1}`
          if (v === null || v === '') return ''
          if (typeof v === 'number') return `<c r="${ref}"><v>${v}</v></c>`
          if (typeof v === 'boolean') return `<c r="${ref}" t="b"><v>${v ? 1 : 0}</v></c>`
          return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeText(v)}</t></is></c>`
        })
        .join('')
      return `<row r="${r + 1}">${cells}</row>`
    })
    .join('')
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`
}

/** Smart Data → a QLCL workbook (`ThongTin` + one sheet per collection, `Data` first). */
export async function exportQlclWorkbook(model: DvhModel): Promise<Uint8Array> {
  const sheets: { name: string; rows: Scalar[][] }[] = [
    {
      name: 'ThongTin',
      rows: [['Tên', 'Giá trị'], ...model.fields.map((f) => [f.name, f.value] as Scalar[])],
    },
    ...[...model.collections]
      .sort((a, b) => (a.name === 'Data' ? -1 : b.name === 'Data' ? 1 : 0))
      .map((c) => ({
        name: c.name.slice(0, 31),
        rows: [c.columns.map((col) => col.title), ...c.rows.map((r) => [...r])],
      })),
  ]
  const zip = new JSZip()
  const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
  zip.file(
    '[Content_Types].xml',
    `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      sheets
        .map(
          (_, i) =>
            `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
        )
        .join('') +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `${X}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
      sheets
        .map(
          (s, i) =>
            `<sheet name="${escapeText(s.name).replace(/"/g, '&quot;')}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`,
        )
        .join('') +
      '</sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      sheets
        .map(
          (_, i) =>
            `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
        )
        .join('') +
      '</Relationships>',
  )
  sheets.forEach((s, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s.rows)))
  return deterministicZip(zip)
}

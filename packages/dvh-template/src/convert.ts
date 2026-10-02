/**
 * DVH-Tool template conversion (P6). Templates of the DVH-Tool / QLCL-DVH
 * add-ins mark data with plain text:
 *
 * - `<<Field name>>` anywhere in a paragraph → a Smart Field bound to the model;
 * - a region between a `BD_Bang[_Name]` marker and a `KT_Bang[_Name]` marker →
 *   a repeating section over a new collection `Name`, its `<<Column>>`
 *   placeholders becoming column controls. In a table the marker rows go and
 *   the rows between them repeat (the table keeps its own formatting); in body
 *   text the marker paragraphs go and the paragraphs between them repeat.
 *
 * The model part is created or extended with the new fields and collections.
 * Placeholders split across several runs (Word does that after edits) are
 * found on the paragraph's text; one that spans a tab, a break or a picture
 * is left as text and reported.
 */
import JSZip from 'jszip'
import {
  emptyModel,
  fieldSdtPrXml,
  itemPropsXml,
  MODEL_NS,
  newDvhId,
  newStoreItemId,
  parseModelXml,
  serializeModelXml,
  storeItemIdOf,
  type DvhCollection,
  type DvhModel,
} from '@genoffice/dvh-model'
import { deterministicZip } from './fill'
import { elements, escapeText, replaceSpans } from './ooxml'
import { findItemByNamespace } from './release'

export interface ConversionReport {
  readonly fields: readonly string[]
  readonly regions: readonly { name: string; columns: readonly string[]; kind: 'rows' | 'blocks' }[]
  readonly warnings: readonly string[]
}

export interface ConversionResult {
  readonly bytes: Uint8Array
  readonly model: DvhModel
  readonly report: ConversionReport
}

const PLACEHOLDER = /<<\s*([^<>]+?)\s*>>/g
/** stands for a non-text run in a paragraph's text: no placeholder may cross it */
const SEPARATOR = String.fromCharCode(0)
const textOf = (xml: string) =>
  [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => decodeEntities(m[1]!)).join('')
const decodeEntities = (s: string) =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')

/** `Khối lượng` → `khoi_luong`: a column key safe in expressions (`row.khoi_luong`). */
export function keyOf(title: string): string {
  const key = title
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
  return /^[a-z_]/.test(key) ? key : `c_${key || 'col'}`
}

interface RunInfo {
  readonly start: number
  readonly end: number
  readonly xml: string
  readonly rPr: string
  /** the run's text, or null for runs holding anything but text */
  readonly text: string | null
}

function runsOf(pXml: string): RunInfo[] {
  return elements(pXml, 'w:r').map((r) => {
    const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/.exec(r.xml)?.[0] ?? ''
    const body = r.xml
      .replace(/^<w:r(?:\s[^>]*)?>/, '')
      .replace(/<\/w:r>$/, '')
      .replace(rPr, '')
    const onlyText = /^(?:\s*<w:t(?:\s[^>]*)?>[^<]*<\/w:t>|\s*<w:t\/>)*\s*$/.test(body)
    return { ...r, rPr, text: onlyText ? textOf(body) : null }
  })
}

const textRun = (rPr: string, text: string) =>
  text ? `<w:r>${rPr}<w:t xml:space="preserve">${escapeText(text)}</w:t></w:r>` : ''

/**
 * Replaces every `<<name>>` of a paragraph with `make(name, rPr)`, whatever
 * runs the placeholder is split over. Returns the paragraph and the names it
 * could not convert (spanning a non-text run).
 */
export function replacePlaceholders(
  pXml: string,
  make: (name: string, rPr: string) => string,
): { xml: string; skipped: string[] } {
  const runs = runsOf(pXml)
  // full text, with a separator for non-text runs so no placeholder crosses them
  const full = runs.map((run) => run.text ?? SEPARATOR).join('')
  const matches = [...full.matchAll(PLACEHOLDER)].map((m) => ({
    name: m[1]!.trim(),
    start: m.index!,
    end: m.index! + m[0].length,
  }))
  const skipped = matches
    .filter((m) => full.slice(m.start, m.end).includes(SEPARATOR))
    .map((m) => m.name.split(SEPARATOR).join(' '))
  const usable = matches.filter((m) => !full.slice(m.start, m.end).includes(SEPARATOR))
  if (usable.length === 0) return { xml: pXml, skipped }
  const spans: { start: number; end: number; with: string }[] = []
  let cursor = 0
  runs.forEach((run) => {
    const length = (run.text ?? SEPARATOR).length
    const from = cursor
    const to = cursor + length
    cursor = to
    if (run.text === null) return
    const touching = usable.filter((m) => m.start < to && m.end > from)
    if (touching.length === 0) return
    let out = ''
    let pos = from
    for (const m of touching) {
      if (m.start > pos) out += textRun(run.rPr, full.slice(pos, m.start))
      // the control goes where the placeholder starts, with that run's formatting
      if (m.start >= from) out += make(m.name, run.rPr)
      pos = Math.min(to, m.end)
    }
    if (pos < to) out += textRun(run.rPr, full.slice(pos, to))
    spans.push({ start: run.start, end: run.end, with: out })
  })
  return { xml: replaceSpans(pXml, spans), skipped }
}

const sdtInline = (tag: string, rPr: string, shown: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${textRun(rPr, shown)}</w:sdtContent></w:sdt>`

const markerName = (text: string, marker: 'BD_Bang' | 'KT_Bang') => {
  const m = new RegExp(`^\\s*${marker}(?:[_\\s-]*([\\p{L}\\p{N}_ ]+))?`, 'u').exec(text)
  return m ? (m[1]?.trim() ?? '') : null
}

/** Converts a DVH-Tool template (docx bytes) into a Smart Template. */
export async function convertDvhToolTemplate(bytes: Uint8Array): Promise<ConversionResult> {
  const zip = await JSZip.loadAsync(bytes)
  const modelPath = await findItemByNamespace(zip, MODEL_NS)
  const model: DvhModel = modelPath
    ? parseModelXml(await zip.file(modelPath)!.async('string'))
    : emptyModel(newDvhId('doc'))
  let storeItemId = modelPath
    ? storeItemIdOf(
        (await zip
          .file(modelPath.replace(/item(\d+)\.xml$/, 'itemProps$1.xml'))
          ?.async('string')) ?? '',
      )
    : null
  storeItemId ??= newStoreItemId()
  const warnings: string[] = []
  const regions: { name: string; columns: string[]; kind: 'rows' | 'blocks' }[] = []
  const fieldNames: string[] = []
  let sdtId = 1000

  const fieldControl = (name: string, rPr: string) => {
    let field = model.fields.find((f) => f.name === name)
    if (!field) {
      field = { id: newDvhId('f'), name, type: 'text', value: null, access: 'readwrite' }
      model.fields.push(field)
      fieldNames.push(name)
    }
    const pr = fieldSdtPrXml({
      fieldId: field.id,
      alias: name,
      sdtId: sdtId++,
      storeItemId: storeItemId!,
    })
    return `<w:sdt>${pr}<w:sdtContent>${textRun(rPr, name)}</w:sdtContent></w:sdt>`
  }

  const newCollection = (name: string): DvhCollection => {
    const unique = (base: string) => {
      let n = 1
      let candidate = base
      while (model.collections.some((c) => c.name === candidate)) candidate = `${base}${++n}`
      return candidate
    }
    const collection: DvhCollection = {
      id: newDvhId('c'),
      name: unique(name || `Bang${regions.length + 1}`),
      columns: [],
      rows: [],
    }
    model.collections.push(collection)
    return collection
  }

  const columnControl = (collection: DvhCollection) => (title: string, rPr: string) => {
    let column = collection.columns.find((c) => c.title === title)
    if (!column) {
      column = { id: newDvhId('c').replace(/^c_/, 'col_'), key: keyOf(title), title, type: 'text' }
      collection.columns.push(column)
    }
    return sdtInline(`dvh:c:${column.id}`, rPr, title)
  }

  const convertParagraphs = (xml: string, make: (name: string, rPr: string) => string) =>
    replaceSpans(
      xml,
      elements(xml, 'w:p').map((p) => {
        const result = replacePlaceholders(p.xml, make)
        for (const name of result.skipped)
          warnings.push(`<<${name}>> spans a tab, break or picture: left as text`)
        return { start: p.start, end: p.end, with: result.xml }
      }),
    )

  const convertTables = (xml: string): string =>
    replaceSpans(
      xml,
      elements(xml, 'w:tbl').flatMap((tbl) => {
        const rows = elements(tbl.xml, 'w:tr')
        const begin = rows.findIndex((r) => markerName(textOf(r.xml), 'BD_Bang') !== null)
        const end = rows.findIndex(
          (r, i) => i > begin && markerName(textOf(r.xml), 'KT_Bang') !== null,
        )
        if (begin < 0 || end < 0) return []
        const collection = newCollection(markerName(textOf(rows[begin]!.xml), 'BD_Bang')!)
        const body = rows
          .slice(begin + 1, end)
          .map((r) => convertParagraphs(r.xml, columnControl(collection)))
          .join('')
        const table =
          tbl.xml.slice(0, rows[0]!.start) +
          rows
            .slice(0, begin)
            .map((r) => r.xml)
            .join('') +
          body +
          rows
            .slice(end + 1)
            .map((r) => r.xml)
            .join('') +
          tbl.xml.slice(rows.at(-1)!.end)
        regions.push({
          name: collection.name,
          columns: collection.columns.map((c) => c.title),
          kind: 'rows',
        })
        const wrapped = `<w:sdt><w:sdtPr><w:alias w:val="${escapeText(collection.name)}"/><w:tag w:val="dvh:repeatrows:${collection.id}"/></w:sdtPr><w:sdtContent>${table}</w:sdtContent></w:sdt>`
        return [{ start: tbl.start, end: tbl.end, with: wrapped }]
      }),
    )

  /** `BD_Bang` … `KT_Bang` paragraphs in body text: the blocks between repeat */
  const convertBlockRegions = (xml: string): string => {
    const body = /<w:body>([\s\S]*)<\/w:body>/.exec(xml)
    if (!body) return xml
    const offset = body.index + '<w:body>'.length
    // top-level blocks only: paragraphs inside tables or controls belong to them
    const candidates = [
      ...elements(body[1]!, 'w:p'),
      ...elements(body[1]!, 'w:tbl'),
      ...elements(body[1]!, 'w:sdt'),
    ]
    const blocks = candidates
      .filter((b) => !candidates.some((o) => o !== b && o.start <= b.start && o.end >= b.end))
      .sort((a, b) => a.start - b.start)
    const spans: { start: number; end: number; with: string }[] = []
    for (let i = 0; i < blocks.length; i++) {
      const name = blocks[i]!.xml.startsWith('<w:p')
        ? markerName(textOf(blocks[i]!.xml), 'BD_Bang')
        : null
      if (name === null) continue
      const j = blocks.findIndex(
        (b, k) =>
          k > i && b.xml.startsWith('<w:p') && markerName(textOf(b.xml), 'KT_Bang') !== null,
      )
      if (j < 0) {
        warnings.push(`BD_Bang${name ? `_${name}` : ''} has no KT_Bang: left as text`)
        continue
      }
      const collection = newCollection(name)
      const inner = blocks
        .slice(i + 1, j)
        .map((b) => convertParagraphs(b.xml, columnControl(collection)))
        .join('')
      regions.push({
        name: collection.name,
        columns: collection.columns.map((c) => c.title),
        kind: 'blocks',
      })
      spans.push({
        start: offset + blocks[i]!.start,
        end: offset + blocks[j]!.end,
        with: `<w:sdt><w:sdtPr><w:alias w:val="${escapeText(collection.name)}"/><w:tag w:val="dvh:repeat:${collection.id}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`,
      })
      i = j
    }
    return replaceSpans(xml, spans)
  }

  for (const path of Object.keys(zip.files).sort()) {
    if (!/^word\/(document|header\d*|footer\d*)\.xml$/.test(path)) continue
    let xml = await zip.file(path)!.async('string')
    xml = convertTables(xml)
    if (path === 'word/document.xml') xml = convertBlockRegions(xml)
    // what is left: fields
    xml = convertParagraphs(xml, fieldControl)
    zip.file(path, xml)
  }

  await writeModelPart(zip, model, modelPath, storeItemId)
  return {
    bytes: await deterministicZip(zip),
    model,
    report: { fields: fieldNames, regions, warnings },
  }
}

/** Writes the model part, creating it (item, props, relationships, content type) when missing. */
async function writeModelPart(
  zip: JSZip,
  model: DvhModel,
  existing: string | null,
  storeItemId: string,
): Promise<void> {
  if (existing) {
    zip.file(existing, serializeModelXml(model))
    return
  }
  let n = 1
  while (zip.file(`customXml/item${n}.xml`)) n++
  zip.file(`customXml/item${n}.xml`, serializeModelXml(model))
  zip.file(`customXml/itemProps${n}.xml`, itemPropsXml(storeItemId, MODEL_NS))
  zip.file(
    `customXml/_rels/item${n}.xml.rels`,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps${n}.xml"/></Relationships>`,
  )
  const relsPath = 'word/_rels/document.xml.rels'
  const rels =
    (await zip.file(relsPath)?.async('string')) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let id = 1
  while (rels.includes(`Id="rIdDvh${id}"`)) id++
  zip.file(
    relsPath,
    rels.replace(
      '</Relationships>',
      `<Relationship Id="rIdDvh${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item${n}.xml"/></Relationships>`,
    ),
  )
  const types = (await zip.file('[Content_Types].xml')?.async('string')) ?? ''
  if (types && !types.includes(`/customXml/itemProps${n}.xml`)) {
    zip.file(
      '[Content_Types].xml',
      types.replace(
        '</Types>',
        `<Override PartName="/customXml/itemProps${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>`,
      ),
    )
  }
}

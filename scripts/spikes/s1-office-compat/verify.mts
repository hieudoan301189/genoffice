// Spike S1: checks the files Word/Excel saved (office-roundtrip.ps1) and pushes
// them through DVH Office's own engines (docx-engine parse + patch save,
// xlsx-gateway structural save + declarative defined-names save).
// Usage: tsx verify.mts <outDir>
import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { parseDocx, saveDocx, type SaveBlock } from '@genoffice/docx-engine'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { applyDefinedNamesState } from '@genoffice/xlsx-gateway/gateway/xlsx-defined-names'
import {
  DATE_VALUE,
  DEFINED_FIELD,
  DEFINED_TABLE,
  DOCX_DOC_ID,
  FIELD_DATE,
  FIELD_PROJECT,
  HISTORY_NS,
  MODEL_NS,
  MODEL_STORE_ID,
  PROJECT_VALUE,
  STALE_A,
  STALE_B,
  TABLE_ITEMS,
  WORD_EDIT_VALUE,
  WORK_ITEMS,
  XLSX_DOC_ID,
} from './dvh-ooxml.mts'

const outDir = process.argv[2] ?? join(import.meta.dirname, 'out')
type Status = 'PASS' | 'FAIL' | 'GAP'
const results: { area: string; check: string; status: Status; detail?: string }[] = []
function check(
  area: string,
  name: string,
  ok: boolean,
  detail?: string,
  onFail: Status = 'FAIL',
): void {
  results.push({ area, check: name, status: ok ? 'PASS' : onFail, ...(detail ? { detail } : {}) })
}

const decode = (s: string): string =>
  s
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')

async function load(name: string): Promise<{ bytes: Buffer; zip: JSZip }> {
  const bytes = readFileSync(join(outDir, name))
  return { bytes, zip: await JSZip.loadAsync(bytes) }
}

/** customXml/itemN.xml whose root lives in `ns`, wherever the writer numbered it */
async function partByNs(zip: JSZip, ns: string): Promise<{ path: string; xml: string } | null> {
  for (const path of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(path)) continue
    const xml = await zip.file(path)!.async('string')
    if (xml.includes(`"${ns}"`)) return { path, xml }
  }
  return null
}

function fieldValue(modelXml: string, id: string): string | null {
  const m = new RegExp(`<(?:\\w+:)?f\\b[^>]*\\bid="${id}"[^>]*>([^<]*)<`).exec(modelXml)
  return m ? decode(m[1]!) : null
}

function objectsJson(modelXml: string): unknown {
  const m = /<(?:\w+:)?objects\b[^>]*>([\s\S]*?)<\/(?:\w+:)?objects>/.exec(modelXml)
  return m ? JSON.parse(decode(m[1]!)) : null
}

function historyCount(historyXml: string): number {
  const m = /<(?:\w+:)?changes\b[^>]*>([\s\S]*?)<\/(?:\w+:)?changes>/.exec(historyXml)
  if (!m) return -1
  return decode(m[1]!)
    .split('\n')
    .filter((line) => JSON.parse(line).id).length
}

async function customXmlEntries(zip: JSZip): Promise<Map<string, string>> {
  const out = new Map<string, string>()
  for (const path of Object.keys(zip.files)) {
    if (path.startsWith('customXml/') && !zip.files[path]!.dir) {
      out.set(path, await zip.file(path)!.async('string'))
    }
  }
  return out
}

function sameEntries(a: Map<string, string>, b: Map<string, string>): boolean {
  if (a.size !== b.size) return false
  for (const [k, v] of a) if (b.get(k) !== v) return false
  return true
}

const blockText = (b: { runs?: { text: string }[] }): string =>
  (b.runs ?? []).map((r) => r.text).join('')

// ---------------- DOCX saved by Word ----------------
async function verifyWordDocx(): Promise<void> {
  const area = 'docx/Word'
  const { bytes, zip } = await load('word-saved-dvh-s1.docx')
  const model = await partByNs(zip, MODEL_NS)
  const history = await partByNs(zip, HISTORY_NS)
  check(area, 'model customXml part kept', !!model, model?.path)
  check(area, 'history customXml part kept', !!history, history?.path)
  if (model) {
    check(area, 'docId kept', model.xml.includes(`docId="${DOCX_DOC_ID}"`))
    check(
      area,
      'Word wrote the content-control edit back into the store (two-way binding)',
      fieldValue(model.xml, FIELD_PROJECT) === WORD_EDIT_VALUE,
      fieldValue(model.xml, FIELD_PROJECT) ?? 'missing',
    )
    check(area, 'untouched field value kept', fieldValue(model.xml, FIELD_DATE) === DATE_VALUE)
    const objects = objectsJson(model.xml) as { tables?: { id: string; rows: unknown }[] } | null
    check(
      area,
      'objects JSON (CDATA) survives',
      objects?.tables?.[0]?.id === TABLE_ITEMS &&
        JSON.stringify(objects.tables[0].rows) === JSON.stringify(WORK_ITEMS),
      model.xml.includes('<![CDATA[') ? 'still CDATA' : 'rewritten as escaped text',
    )
  }
  const docXml = await zip.file('word/document.xml')!.async('string')
  check(area, 'field SDT tag kept', docXml.includes(`w:tag w:val="dvh:f:${FIELD_PROJECT}"`))
  check(
    area,
    'dataBinding kept (storeItemID)',
    docXml.includes(`w:storeItemID="${MODEL_STORE_ID}"`),
  )
  check(area, 'table SDT tag kept', docXml.includes(`w:tag w:val="dvh:t:${TABLE_ITEMS}"`))
  check(
    area,
    'Word refreshed the stale content from the store',
    !docXml.includes(STALE_A) && !docXml.includes(STALE_B),
  )
  check(area, 'paragraph added in Word present', docXml.includes('Paragraph added in Word.'))

  // DVH Office reads the Word-saved file
  const parsed = await parseDocx(bytes)
  const visible = parsed.blocks.filter((b) => !b.hidden)
  const fieldBlock = visible.find((b) => blockText(b).includes(WORD_EDIT_VALUE))
  check(
    'docx/DVH',
    'docx-engine parses the Word-saved file and shows the bound value',
    !!fieldBlock,
  )
  check(
    'docx/DVH',
    'table inside the block SDT parses as a table',
    visible.some((b) => b.type === 'table'),
  )

  // DVH Office edits an unrelated paragraph and saves (paragraph-patch save)
  const plain = visible.find((b) => blockText(b).startsWith('Đoạn thường'))
  const generated = (text: string): SaveBlock => ({
    kind: 'generated',
    block: { type: 'paragraph', runs: [{ text, bold: false }] },
  })
  const saved = await saveDocx(
    parsed,
    visible.map((b) =>
      b === plain
        ? generated('DVH Office đã sửa đoạn này.')
        : { kind: 'original', docxIndex: b.docxIndex! },
    ),
  )
  const savedZip = await JSZip.loadAsync(saved)
  check(
    'docx/DVH',
    'save after editing another paragraph keeps every customXml entry byte-identical',
    sameEntries(await customXmlEntries(zip), await customXmlEntries(savedZip)),
  )
  const savedDoc = await savedZip.file('word/document.xml')!.async('string')
  check(
    'docx/DVH',
    'save keeps the untouched inline field SDT + dataBinding',
    savedDoc.includes(`w:tag w:val="dvh:f:${FIELD_PROJECT}"`) &&
      savedDoc.includes(`w:storeItemID="${MODEL_STORE_ID}"`),
  )
  check(
    'docx/DVH',
    'save keeps the table SDT',
    savedDoc.includes(`w:tag w:val="dvh:t:${TABLE_ITEMS}"`),
  )

  // DVH Office edits the paragraph that holds the field: today the run-level SDT is flattened
  const fieldPara = visible.find((b) => blockText(b).startsWith('Tên dự án'))!
  const rewritten = await saveDocx(
    parsed,
    visible.map((b) =>
      b === fieldPara
        ? generated(`Tên dự án: ${WORD_EDIT_VALUE} (sửa trong DVH Office)`)
        : { kind: 'original', docxIndex: b.docxIndex! },
    ),
  )
  const rewrittenDoc = await (
    await JSZip.loadAsync(rewritten)
  )
    .file('word/document.xml')!
    .async('string')
  check(
    'docx/DVH',
    'editing the paragraph that holds a field keeps the field SDT',
    rewrittenDoc.includes(`w:tag w:val="dvh:f:${FIELD_PROJECT}"`),
    'run-level w:sdt is flattened when its paragraph regenerates — spike S2 / P1 must add an inline field node',
    'GAP',
  )
}

// ---------------- XLSX saved by Excel ----------------
function definedNames(workbookXml: string): Map<string, { formula: string; hidden: boolean }> {
  const out = new Map<string, { formula: string; hidden: boolean }>()
  for (const m of workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const name = /\bname="([^"]*)"/.exec(m[1]!)?.[1] ?? ''
    out.set(decode(name), { formula: decode(m[2]!), hidden: /\bhidden="(?:1|true)"/.test(m[1]!) })
  }
  return out
}

async function verifyExcelXlsx(): Promise<void> {
  const area = 'xlsx/Excel'
  const { bytes, zip } = await load('excel-saved-dvh-s1.xlsx')
  const model = await partByNs(zip, MODEL_NS)
  const history = await partByNs(zip, HISTORY_NS)
  check(area, 'model customXml part kept', !!model, model?.path)
  check(area, 'history customXml part kept', !!history, history?.path)
  if (model) {
    check(
      area,
      'docId + field value kept',
      model.xml.includes(`docId="${XLSX_DOC_ID}"`) &&
        fieldValue(model.xml, FIELD_PROJECT) === PROJECT_VALUE,
    )
    const objects = objectsJson(model.xml) as { tables?: { id: string }[] } | null
    check(area, 'objects JSON (CDATA) survives', objects?.tables?.[0]?.id === TABLE_ITEMS)
  }
  if (history) {
    const original = await partByNs((await load('dvh-s1.xlsx')).zip, HISTORY_NS)
    check(
      area,
      'history entries intact',
      historyCount(history.xml) === historyCount(original!.xml),
      `${historyCount(history.xml)} entries`,
    )
  }
  const workbookXml = await zip.file('xl/workbook.xml')!.async('string')
  const names = definedNames(workbookXml)
  const field = names.get(DEFINED_FIELD)
  const table = names.get(DEFINED_TABLE)
  check(
    area,
    'hidden field name kept hidden and moved B5→B6 with the inserted row',
    field?.hidden === true && field.formula === 'Sheet1!$B$6',
    JSON.stringify(field),
  )
  check(
    area,
    'hidden table name moved A8:C11→A9:C12',
    table?.hidden === true && table.formula === 'Sheet1!$A$9:$C$12',
    JSON.stringify(table),
  )
  check(area, 'visible control name moved too', names.get('TenDuAn')?.formula === 'Sheet1!$B$6')

  // DVH Office structural save on the Excel-saved file: insert 2 rows at the top
  const mutation = await applyCellEditsToXlsx(
    bytes,
    [],
    [{ sheetName: 'Sheet1', ops: [{ kind: 'insert-rows', index: 0, count: 2 }] }],
  )
  const after = await JSZip.loadAsync(mutation.buffer)
  const shifted = definedNames(await after.file('xl/workbook.xml')!.async('string'))
  check(
    'xlsx/DVH',
    'DVH Office save shifts hidden _dvh names with inserted rows (B6→B8)',
    shifted.get(DEFINED_FIELD)?.formula === 'Sheet1!$B$8' &&
      shifted.get(DEFINED_FIELD)?.hidden === true,
    JSON.stringify(shifted.get(DEFINED_FIELD)),
  )
  check(
    'xlsx/DVH',
    'DVH Office structural save keeps every customXml entry byte-identical',
    sameEntries(await customXmlEntries(zip), await customXmlEntries(after)),
  )

  // Name Manager save (declarative rewrite) must leave hidden _dvh names alone
  const rewritten = applyDefinedNamesState(workbookXml, {
    names: [
      { name: 'TenDuAn', formula: 'Sheet1!$B$6' },
      { name: 'NewName', formula: 'Sheet1!$A$1' },
    ],
    preserveNames: [],
  })
  const kept = definedNames(rewritten)
  check(
    'xlsx/DVH',
    'Name Manager save keeps hidden _dvh names verbatim and does not duplicate them',
    kept.get(DEFINED_FIELD)?.hidden === true &&
      kept.get(DEFINED_TABLE)?.hidden === true &&
      (rewritten.match(new RegExp(DEFINED_FIELD.replace(/\./g, '\\.'), 'g')) ?? []).length === 1,
  )
}

// ---------------- size probe ----------------
async function verifyBig(): Promise<void> {
  for (const [name, area] of [
    ['word-saved-dvh-s1-big.docx', 'docx/Word'],
    ['excel-saved-dvh-s1-big.xlsx', 'xlsx/Excel'],
  ] as const) {
    const { bytes, zip } = await load(name)
    const history = await partByNs(zip, HISTORY_NS)
    check(
      area,
      `5 MB history part kept (${name})`,
      !!history && history.xml.length > 4_900_000,
      `${history?.xml.length ?? 0} chars, file ${bytes.length} bytes`,
    )
  }
}

await verifyWordDocx()
await verifyExcelXlsx()
await verifyBig()

for (const r of results)
  console.log(`${r.status.padEnd(4)} [${r.area}] ${r.check}${r.detail ? ` — ${r.detail}` : ''}`)
const fail = results.filter((r) => r.status === 'FAIL').length
console.log(
  `\n${results.length} checks: ${results.filter((r) => r.status === 'PASS').length} pass, ${fail} fail, ${results.filter((r) => r.status === 'GAP').length} known gap`,
)
writeFileSync(join(outDir, 'verify-report.json'), JSON.stringify(results, null, 2))
process.exitCode = fail ? 1 : 0

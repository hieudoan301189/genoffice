/**
 * Field write-back into a closed workbook (P3 Read/Write links).
 *
 * A document writes a linked field back to its source workbook when that
 * workbook is not open in Sheets. Three places change, nothing else:
 *
 * - the cell the hidden name `_dvh.f.<id>` points at (what Excel shows);
 * - the field in the model part, text and revision, patched in place with
 *   setFieldTextInXml so every other byte stays as the last writer left it;
 * - the history part, which gains one `Data.SetField` change set.
 *
 * Before writing, each field is checked against the base the document edited
 * from (revision + text, the text read from the cell): a workbook changed in
 * the meantime, by DVH or by Excel, reports a conflict instead.
 */
import JSZip from 'jszip'
import {
  checkFieldWrite,
  fieldText,
  FIELD_NAME_PREFIX,
  HISTORY_NS,
  MODEL_NS,
  modelHash,
  newDvhId,
  newStoreItemId,
  parseHistoryXml,
  parseModelXml,
  serializeHistoryXml,
  setFieldTextInXml,
  type ChangeSet,
  type FieldWrite,
  type FieldWriteResult,
  type HistoryPart,
  type Scalar,
} from '@genoffice/dvh-model'
import { findCustomXmlItemByNamespace } from '@genoffice/docx-engine'
import { applyDvhCustomXmlParts } from '@genoffice/xlsx-gateway/gateway/xlsx-dvh'
import { columnIndex, modelFromCells, parseRef, workbookLayout } from './dvh-xlsx-cells'

export interface XlsxFieldWrite {
  /** the patched workbook, null when nothing was written */
  readonly bytes: Uint8Array | null
  readonly results: FieldWriteResult[]
}

const escapeText = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

const columnLetters = (index: number): string => {
  let n = index + 1
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

/** The cell element for a value, keeping the old cell's style. */
function cellXml(ref: string, style: string | undefined, value: Scalar): string {
  const s = style ? ` s="${style}"` : ''
  if (value === null || value === '') return `<c r="${ref}"${s}/>`
  if (typeof value === 'number') return `<c r="${ref}"${s}><v>${value}</v></c>`
  if (typeof value === 'boolean') return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${escapeText(value)}</t></is></c>`
}

const CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g
const cellRefOf = (attrs: string) => /\sr="([A-Z]{1,3})(\d+)"/.exec(attrs)

/**
 * Sets one cell in a worksheet part. Returns null when the cell holds a
 * formula (a computed value is not written over).
 */
export function setCellInSheetXml(
  sheetXml: string,
  row: number,
  column: number,
  value: Scalar,
): string | null {
  const ref = `${columnLetters(column)}${row + 1}`
  // the cell itself
  for (const m of sheetXml.matchAll(CELL_RE)) {
    const pos = cellRefOf(m[1] ?? '')
    if (!pos || `${pos[1]}${pos[2]}` !== ref) continue
    if (/<f\b/.test(m[2] ?? '')) return null
    const style = /\ss="(\d+)"/.exec(m[1] ?? '')?.[1]
    return (
      sheetXml.slice(0, m.index) +
      cellXml(ref, style, value) +
      sheetXml.slice(m.index! + m[0].length)
    )
  }
  const cell = cellXml(ref, undefined, value)
  // its row: insert in column order
  const rowRe = new RegExp(`<row\\b[^>]*\\sr="${row + 1}"[^>]*?(?:/>|>([\\s\\S]*?)</row>)`)
  const rowMatch = rowRe.exec(sheetXml)
  if (rowMatch) {
    const whole = rowMatch[0]
    if (whole.endsWith('/>')) {
      const replaced = `${whole.slice(0, -2)}>${cell}</row>`
      return (
        sheetXml.slice(0, rowMatch.index) + replaced + sheetXml.slice(rowMatch.index + whole.length)
      )
    }
    const bodyStart = whole.indexOf('>') + 1
    const body = rowMatch[1] ?? ''
    let insertAt = body.length
    for (const m of body.matchAll(CELL_RE)) {
      const pos = cellRefOf(m[1] ?? '')
      if (pos && columnIndex(pos[1]!) > column) {
        insertAt = m.index!
        break
      }
    }
    const nextBody = body.slice(0, insertAt) + cell + body.slice(insertAt)
    const replaced = whole.slice(0, bodyStart) + nextBody + whole.slice(bodyStart + body.length)
    return (
      sheetXml.slice(0, rowMatch.index) + replaced + sheetXml.slice(rowMatch.index + whole.length)
    )
  }
  // no row yet: insert in row order
  const newRow = `<row r="${row + 1}">${cell}</row>`
  if (/<sheetData\s*\/>/.test(sheetXml)) {
    return sheetXml.replace(/<sheetData\s*\/>/, `<sheetData>${newRow}</sheetData>`)
  }
  for (const m of sheetXml.matchAll(/<row\b[^>]*\sr="(\d+)"/g)) {
    if (Number(m[1]) > row + 1)
      return sheetXml.slice(0, m.index) + newRow + sheetXml.slice(m.index!)
  }
  return sheetXml.replace('</sheetData>', `${newRow}</sheetData>`)
}

/** Asks Excel to recompute on open: formulas that read a written cell keep stale cached values otherwise. */
function withFullCalcOnLoad(workbookXml: string): string {
  if (/<calcPr\b[^>]*fullCalcOnLoad="1"/.test(workbookXml)) return workbookXml
  if (/<calcPr\b/.test(workbookXml)) {
    return workbookXml.replace(
      /<calcPr\b([^>]*?)(\s*\/?>)/,
      (_, attrs: string, end: string) =>
        `<calcPr${attrs.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${end}`,
    )
  }
  return workbookXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
}

/**
 * Writes fields into a workbook's bytes. Throws when the workbook is not the
 * linked one (other docId) or has no DVH model.
 */
export async function writeFieldsToXlsx(
  bytes: Uint8Array,
  request: { docId: string; writes: readonly FieldWrite[] },
): Promise<XlsxFieldWrite> {
  const zip = await JSZip.loadAsync(bytes)
  const modelPath = await findCustomXmlItemByNamespace(zip, MODEL_NS)
  if (!modelPath) throw new Error('The workbook has no DVH Smart Data.')
  let modelXml = await zip.file(modelPath)!.async('string')
  const model = parseModelXml(modelXml)
  if (model.docId !== request.docId) throw new Error('The workbook is not the linked source.')
  // the cells are the truth for "current": Excel edits never reach the model part
  const fromCells = await modelFromCells(bytes, model)
  const layout = await workbookLayout(zip)
  const sheetXmls = new Map<string, string>()
  const results: FieldWriteResult[] = []
  const changes: ChangeSet['changes'] = []

  for (const write of request.writes) {
    const field = model.fields.find((f) => f.id === write.fieldId)
    const cellField = fromCells.fields.find((f) => f.id === write.fieldId)
    if (!field || !cellField) {
      results.push({ fieldId: write.fieldId, status: 'missing' })
      continue
    }
    const current = { rev: field.rev ?? 0, text: fieldText(cellField.value) }
    if (checkFieldWrite(write, current) === 'conflict') {
      results.push({ fieldId: write.fieldId, status: 'conflict', current })
      continue
    }
    const rect = parseRef(layout.names.get(`${FIELD_NAME_PREFIX}${field.id}`) ?? '')
    const sheetPath = rect ? layout.sheetPaths.get(rect.sheetName) : undefined
    if (rect && sheetPath) {
      const sheetXml = sheetXmls.get(sheetPath) ?? (await zip.file(sheetPath)?.async('string'))
      if (sheetXml === undefined) {
        results.push({ fieldId: write.fieldId, status: 'missing', current })
        continue
      }
      const next = setCellInSheetXml(sheetXml, rect.row, rect.column, write.value)
      if (next === null) {
        results.push({ fieldId: write.fieldId, status: 'formula', current })
        continue
      }
      sheetXmls.set(sheetPath, next)
    }
    // a field whose name is gone keeps only its model value, as in Sheets
    const rev = (field.rev ?? 0) + 1
    const text = fieldText(write.value)
    const patched = setFieldTextInXml(modelXml, field.id, text, rev)
    if (patched === null) {
      results.push({ fieldId: write.fieldId, status: 'missing', current })
      continue
    }
    modelXml = patched
    changes.push({ objectId: field.id, path: 'value', before: cellField.value, after: write.value })
    results.push({ fieldId: write.fieldId, status: 'written', current: { rev, text } })
  }

  if (changes.length === 0) return { bytes: null, results }

  for (const [path, xml] of sheetXmls) zip.file(path, xml)
  zip.file(modelPath, modelXml)
  zip.file('xl/workbook.xml', withFullCalcOnLoad(layout.workbookXml))

  const changeSet: ChangeSet = {
    id: newDvhId('r').replace(/^r_/, 'cs_'),
    txId: newDvhId('r').replace(/^r_/, 'tx_'),
    docId: model.docId,
    at: new Date().toISOString(),
    source: 'link',
    action: 'Data.SetField',
    changes,
  }
  const historyPath = await findCustomXmlItemByNamespace(zip, HISTORY_NS)
  let history: HistoryPart = { docId: model.docId, changes: [] }
  if (historyPath) {
    try {
      history = parseHistoryXml(await zip.file(historyPath)!.async('string'))
    } catch {
      // an unreadable history part is replaced rather than blocking the write
    }
  }
  // the written model is a DVH save: it becomes the base the next open compares with
  const written = parseModelXml(modelXml)
  const nextHistory: HistoryPart = {
    ...history,
    docId: model.docId,
    modelHash: modelHash(written),
    base: written,
    changes: [...history.changes, changeSet],
  }
  const touched = new Set<string>()
  await applyDvhCustomXmlParts(
    {
      paths: async () => Object.keys(zip.files),
      has: async (path) => zip.file(path) !== null,
      readText: async (path) => (await zip.file(path)?.async('string')) ?? '',
      write: (path, content) => void zip.file(path, content),
      add: (path, content) => void zip.file(path, content),
    },
    [{ ns: HISTORY_NS, xml: serializeHistoryXml(nextHistory), storeItemId: newStoreItemId() }],
    touched,
  )
  const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  return { bytes: out, results }
}

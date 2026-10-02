/**
 * P5 acceptance, Docs side:
 * - restore one field or one table (collection) without rolling the file back;
 * - save then reopen keeps the history (and its base);
 * - a Word edit of the model between two DVH saves becomes an `external` change set;
 * - a clean export carries no history (or no DVH at all);
 * - undoing one transaction asks first when later edits touched the same objects;
 * - the history stays within the P0 size limit and compacts on request.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import {
  emptyModel,
  HISTORY_NS,
  MODEL_NS,
  parseHistoryXml,
  serializeModelXml,
  setFieldTextInXml,
  type DvhCollection,
  type DvhModel,
} from '@genoffice/dvh-model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  compactDvhHistory,
  dvhDocsCustomXmlParts,
  fieldOccurrences,
  historyEntries,
  historyStatus,
  insertSmartField,
  linkSource,
  loadDvhDocs,
  setActiveDvhDocs,
  updateFromSource,
  type DvhDocsState,
} from '../src/renderer/dvh-smart-data'
import { createDocTable, insertDocTable, tableOccurrences } from '../src/renderer/dvh-tables'
import { restoreObject, revertTransaction } from '../src/renderer/dvh-history'
import { applyReleasePolicy, unwrapDvhControls } from '../src/main/dvh-release'

const editors: Editor[] = []
afterEach(() => {
  setActiveDvhDocs(null)
  for (const e of editors.splice(0)) e.destroy()
})

const items = (rows: (string | number)[][]): DvhCollection => ({
  id: 'c_items',
  name: 'Items',
  columns: [
    { id: 'k_code', key: 'code', title: 'Mã', type: 'text' },
    { id: 'k_qty', key: 'qty', title: 'KL', type: 'number' },
  ],
  rows,
})
const workbook = (name: string, rows: (string | number)[][]): DvhModel => ({
  ...emptyModel('doc_workbook0000001'),
  fields: [{ id: 'f_name', name: 'Project.Name', type: 'text', value: name, access: 'readwrite' }],
  collections: [items(rows)],
})

async function parsedWith(parts: Record<string, string> = {}): Promise<ParsedDocFull> {
  const zip = new JSZip()
  for (const [path, xml] of Object.entries(parts)) zip.file(path, xml)
  return {
    internal: { originalBytes: await zip.generateAsync({ type: 'uint8array' }) },
  } as unknown as ParsedDocFull
}

function newEditor(): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Dự án: ' }] }],
    } as never,
  })
  editors.push(editor)
  return editor
}

const rowsOf = (editor: Editor) => {
  const out: string[][] = []
  tableOccurrences(editor.state.doc)[0]!.node.forEach((row) => {
    const cells: string[] = []
    row.forEach((cell) => cells.push(cell.textContent))
    out.push(cells)
  })
  return out.slice(1)
}

/** a linked document with a field and a table, pulled three times (v1 → v2 → v3) */
async function linkedDocument(): Promise<{ editor: Editor; dvh: DvhDocsState; pulls: string[] }> {
  const dvh = await loadDvhDocs(await parsedWith())
  setActiveDvhDocs(dvh)
  const editor = newEditor()
  const link = linkSource(dvh, { path: 'D:\\w\\kl.xlsx', model: workbook('v1', [['A', 1]]) }, null)
  editor.commands.setTextSelection(editor.state.doc.content.size - 1)
  insertSmartField(editor, dvh, dvh.model!.fields[0]!)
  insertDocTable(editor, dvh, createDocTable(dvh, { collectionId: 'c_items' }), {
    afterBlockIndex: 0,
  })
  const pulls: string[] = []
  for (const [name, rows] of [
    [
      'v2',
      [
        ['A', 2],
        ['B', 5],
      ],
    ],
    ['v3', [['A', 3]]],
  ] as const) {
    updateFromSource(editor, dvh, link, {
      path: 'D:\\w\\kl.xlsx',
      model: workbook(
        name,
        rows.map((r) => [...r]),
      ),
    })
    pulls.push(dvh.pending.at(-1)!.id)
  }
  return { editor, dvh, pulls }
}

describe('object history (P5)', () => {
  it('restores one field and one collection on their own, as new change sets', async () => {
    const { editor, dvh, pulls } = await linkedDocument()
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('v3')
    // the table after v3 shows one row; restore the collection to v2 only
    const refreshTable = () => rowsOf(editor)
    restoreObject(editor, dvh, pulls[0]!, 'c_items')
    expect(dvh.model!.collections[0]!.rows).toEqual([
      ['A', 2],
      ['B', 5],
    ])
    expect(refreshTable()).toEqual([
      ['A', '2'],
      ['B', '5'],
    ])
    // the field is untouched by the table restore
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('v3')
    restoreObject(editor, dvh, pulls[0]!, 'f_name')
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('v2')
    // the restores and the table re-render they caused are all restore change sets
    const restores = dvh.pending.filter((c) => c.source === 'restore')
    expect(restores.map((c) => c.action)).toEqual([
      'History.RestoreObject',
      'Table.Refresh',
      'History.RestoreObject',
    ])
    // nothing was removed from the history
    expect(historyEntries(dvh).some((c) => c.id === pulls[1])).toBe(true)
  })

  it('keeps the history through save and reopen, and records a Word edit as external', async () => {
    const { editor, dvh } = await linkedDocument()
    // the parts a save writes for this document (its state, registered for the parsed file)
    const parsed = await parsedWith()
    Object.assign(await loadDvhDocs(parsed), dvh)
    const saved = dvhDocsCustomXmlParts(parsed, editor)!
    const modelXml = saved.find((p) => p.ns === MODEL_NS)!.xml
    const historyXml = saved.find((p) => p.ns === HISTORY_NS)!.xml
    const history = parseHistoryXml(historyXml)
    expect(history.changes.length).toBe(dvh.pending.length)
    expect(history.base?.fields[0]!.value).toBe('v3')

    const reopened = await loadDvhDocs(
      await parsedWith({ 'customXml/item1.xml': modelXml, 'customXml/item2.xml': historyXml }),
    )
    expect(reopened.history!.changes.map((c) => c.id)).toEqual(history.changes.map((c) => c.id))
    expect(reopened.pending).toEqual([])

    // Word's two-way binding rewrote the field in the model part
    const edited = setFieldTextInXml(modelXml, 'f_name', 'Sửa trong Word')!
    const afterWord = await loadDvhDocs(
      await parsedWith({ 'customXml/item1.xml': edited, 'customXml/item2.xml': historyXml }),
    )
    expect(afterWord.pending).toHaveLength(1)
    expect(afterWord.pending[0]).toMatchObject({
      source: 'external',
      changes: [{ objectId: 'f_name', path: 'value', before: 'v3', after: 'Sửa trong Word' }],
    })
  })

  it('undoes a transaction, asking first when later edits touched the same objects', async () => {
    const { editor, dvh, pulls } = await linkedDocument()
    const firstTx = historyEntries(dvh).find((c) => c.id === pulls[0])!.txId
    const blocked = revertTransaction(editor, dvh, firstTx)
    expect(blocked.results).toEqual([])
    expect(blocked.conflicts.map((c) => c.objectId).sort()).toEqual(['c_items', 'f_name'])
    const forced = revertTransaction(editor, dvh, firstTx, { force: true })
    expect(forced.results.every((r) => r.applied)).toBe(true)
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('v1')
  })

  it('stays within the size limit and compacts on request', async () => {
    const { dvh } = await linkedDocument()
    const status = historyStatus(dvh)
    expect(status.bytes).toBeGreaterThan(0)
    expect(status.bytes).toBeLessThan(status.limit)
    const before = historyEntries(dvh).length
    compactDvhHistory(dvh, 1)
    const after = historyEntries(dvh)
    expect(after.length).toBe(2)
    expect(after.at(-1)!.action).toBe('History.Compact')
    expect(before).toBeGreaterThan(2)
  })
})

describe('release policy (clean export)', () => {
  const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  async function docx(): Promise<Uint8Array> {
    const zip = new JSZip()
    zip.file(
      '[Content_Types].xml',
      `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/customXml/itemProps1.xml" ContentType="x"/><Override PartName="/customXml/itemProps2.xml" ContentType="x"/></Types>`,
    )
    zip.file(
      'word/_rels/document.xml.rels',
      `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId8" Type="t" Target="../customXml/item1.xml"/><Relationship Id="rId9" Type="t" Target="../customXml/item2.xml"/><Relationship Id="rId1" Type="t" Target="styles.xml"/></Relationships>`,
    )
    const field =
      '<w:sdt><w:sdtPr><w:tag w:val="dvh:f:f_name"/></w:sdtPr><w:sdtContent><w:r><w:t>v3</w:t></w:r></w:sdtContent></w:sdt>'
    const other =
      '<w:sdt><w:sdtPr><w:tag w:val="citation"/></w:sdtPr><w:sdtContent><w:r><w:t>[1]</w:t></w:r></w:sdtContent></w:sdt>'
    zip.file(
      'word/document.xml',
      `${X}<w:document xmlns:w="w"><w:body><w:p><w:r><w:t>Dự án: </w:t></w:r>${field}${other}</w:p>` +
        `<w:sdt><w:sdtPr><w:tag w:val="dvh:t:t_1"/></w:sdtPr><w:sdtContent><w:tbl><w:tr><w:tc><w:p>${field}</w:p></w:tc></w:tr></w:tbl></w:sdtContent></w:sdt></w:body></w:document>`,
    )
    zip.file('customXml/item1.xml', serializeModelXml(workbook('v3', [])))
    zip.file('customXml/itemProps1.xml', '<p/>')
    zip.file(
      'customXml/item2.xml',
      `${X}<dvh:history xmlns:dvh="${HISTORY_NS}" docId="doc_x"><dvh:changes><![CDATA[${JSON.stringify(
        {
          id: 'cs_1',
          txId: 'tx_1',
          docId: 'doc_x',
          at: '2026-01-01T00:00:00.000Z',
          source: 'ui',
          action: 'Data.SetField',
          changes: [],
        },
      )}\n${JSON.stringify({
        id: 'cs_2',
        txId: 'tx_2',
        docId: 'doc_x',
        at: '2026-09-01T00:00:00.000Z',
        source: 'ui',
        action: 'Data.SetField',
        changes: [],
      })}]]></dvh:changes></dvh:history>`,
    )
    zip.file('customXml/itemProps2.xml', '<p/>')
    return zip.generateAsync({ type: 'uint8array' })
  }
  const files = async (bytes: Uint8Array) =>
    Object.keys((await JSZip.loadAsync(bytes)).files).sort()
  const text = async (bytes: Uint8Array, path: string) =>
    (await JSZip.loadAsync(bytes)).file(path)!.async('string')

  it('none: no history part, Smart Data kept; from: older change sets dropped', async () => {
    const none = await applyReleasePolicy(await docx(), { kind: 'none' })
    expect(await files(none)).not.toContain('customXml/item2.xml')
    expect(await files(none)).toContain('customXml/item1.xml')
    expect(await text(none, 'word/_rels/document.xml.rels')).not.toContain('item2.xml')
    expect(await text(none, '[Content_Types].xml')).not.toContain('itemProps2.xml')
    const from = await applyReleasePolicy(await docx(), {
      kind: 'from',
      at: '2026-06-01T00:00:00.000Z',
    })
    expect(
      parseHistoryXml(await text(from, 'customXml/item2.xml')).changes.map((c) => c.id),
    ).toEqual(['cs_2'])
  })

  it('strip: no DVH parts, DVH controls unwrapped (text kept), other controls untouched', async () => {
    const strip = await applyReleasePolicy(await docx(), { kind: 'strip' })
    expect((await files(strip)).filter((f) => f.startsWith('customXml/'))).toEqual([])
    const body = await text(strip, 'word/document.xml')
    expect(body).not.toContain('dvh:')
    expect(body).toContain('<w:t>v3</w:t>')
    expect(body).toContain('<w:tag w:val="citation"/>')
    expect(body).toContain('<w:tbl>')
    expect(unwrapDvhControls('<w:p/>')).toBe('<w:p/>')
  })
})

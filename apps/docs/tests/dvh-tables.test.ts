/**
 * P2: DVH.Table in Docs. A linked collection renders as a docx table inside a
 * dvh:t content control, follows the source on refresh, and refuses to
 * overwrite hand edits unless forced.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { emptyModel, type DvhCollection, type DvhModel } from '@genoffice/dvh-model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pmDocToSavePlan, pmTableToModel, type PmNode } from '../src/renderer/editor/convert'
import { linkSource, loadDvhDocs } from '../src/renderer/dvh-smart-data'
import {
  createDocTable,
  insertDocTable,
  refreshDocTable,
  setDocTableStyle,
  tableOccurrences,
  updateLinkFromSource,
} from '../src/renderer/dvh-tables'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function editorWith(...texts: string[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: texts.map((text) => ({ type: 'docParagraph', content: [{ type: 'text', text }] })),
    } as never,
  })
  editors.push(editor)
  return editor
}

async function emptyDvh() {
  const zip = new JSZip()
  const originalBytes = await zip.generateAsync({ type: 'uint8array' })
  return loadDvhDocs({ internal: { originalBytes } } as unknown as ParsedDocFull)
}

const workItems = (rows: (string | number)[][]): DvhCollection => ({
  id: 'c_workitems0000001',
  name: 'WorkItems',
  columns: [
    {
      id: 'k_code',
      key: 'code',
      title: 'Mã',
      type: 'text',
      headerStyle: { bold: true, fill: '#FFF2CC' },
    },
    { id: 'k_qty', key: 'qty', title: 'Khối lượng', type: 'number', numFmt: '#,##0.00' },
  ],
  rows,
})
const workbook = (rows: (string | number)[][]): DvhModel => ({
  ...emptyModel('doc_workbook00000001'),
  collections: [workItems(rows)],
})

const cellTexts = (node: { forEach: (fn: (n: never) => void) => void }) => {
  const out: string[][] = []
  node.forEach((row: { forEach: (fn: (c: { textContent: string }) => void) => void }) => {
    const cells: string[] = []
    row.forEach((cell) => cells.push(cell.textContent))
    out.push(cells)
  })
  return out
}

describe('DVH.Table in Docs', () => {
  it('inserts a linked collection as a table with formats, header style and a repeating header', async () => {
    const dvh = await emptyDvh()
    const editor = editorWith('Bảng khối lượng:', 'Hết.')
    linkSource(
      dvh,
      {
        path: 'D:\\w\\kl.xlsx',
        model: workbook([
          ['AB.1', 120.5],
          ['AB.2', 2450.25],
        ]),
      },
      'D:\\w\\bc.docx',
    )
    editor.commands.setTextSelection(3)
    const table = createDocTable(dvh, { collectionId: 'WorkItems' })
    insertDocTable(editor, dvh, table)

    const [occ] = tableOccurrences(editor.state.doc)
    expect(occ?.tableId).toBe(table.id)
    expect(editor.state.doc.child(1).type.name).toBe('docTable')
    expect(cellTexts(occ!.node as never)).toEqual([
      ['Mã', 'Khối lượng'],
      ['AB.1', '120.50'],
      ['AB.2', '2,450.25'],
    ])
    const model = pmTableToModel(occ!.node.toJSON() as PmNode)
    expect(model.repeatHeaderRows?.[0]).toBe(true)
    expect(model.rows[0]?.[0]?.fill).toBe('FFF2CC')

    // saving wraps the generated table in its dvh:t control
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, [])
    const xml = plan.saveBlocks
      .map((b) => (b.kind === 'xml' ? b.xml : ''))
      .find((x) => x.includes('dvh:t:'))!
    expect(xml.startsWith(`<w:sdt>${occ!.sdtPr}<w:sdtContent><w:tbl>`)).toBe(true)
    expect(xml).toContain('<w:tblHeader/>')
    expect(xml.endsWith('</w:tbl></w:sdtContent></w:sdt>')).toBe(true)
  })

  it('update from source refreshes the table; hand edits need force', async () => {
    const dvh = await emptyDvh()
    const editor = editorWith('A')
    const link = linkSource(dvh, { path: 'D:\\w\\kl.xlsx', model: workbook([['AB.1', 1]]) }, null)
    const table = createDocTable(dvh, { collectionId: 'c_workitems0000001', mode: 'destination' })
    insertDocTable(editor, dvh, table, { afterBlockIndex: 0 })

    const grown = workbook([
      ['AB.1', 1],
      ['AB.2', 2],
      ['AB.3', 3],
    ])
    const update = updateLinkFromSource(editor, dvh, link, { path: 'D:\\w\\kl.xlsx', model: grown })
    expect(update).toMatchObject({ tables: 1, editedTables: [] })
    expect(cellTexts(tableOccurrences(editor.state.doc)[0]!.node as never)).toHaveLength(4)

    // type into a body cell, then refresh
    const occ = tableOccurrences(editor.state.doc)[0]!
    let cellPos = -1
    occ.node.descendants((node, pos) => {
      if (cellPos < 0 && node.isText && node.text === 'AB.2') cellPos = occ.pos + 1 + pos
    })
    editor.view.dispatch(editor.state.tr.insertText('X', cellPos))
    expect(refreshDocTable(editor, dvh, table.id)).toEqual({ refreshed: 0, needsConfirm: true })
    const shrunk = workbook([['AB.1', 1]])
    const second = updateLinkFromSource(editor, dvh, link, {
      path: 'D:\\w\\kl.xlsx',
      model: shrunk,
    })
    expect(second.editedTables).toEqual([table.id])
    expect(refreshDocTable(editor, dvh, table.id, { force: true })).toEqual({
      refreshed: 1,
      needsConfirm: false,
    })
    expect(cellTexts(tableOccurrences(editor.state.doc)[0]!.node as never)).toEqual([
      ['Mã', 'Khối lượng'],
      ['AB.1', '1.00'],
    ])
    // the destination style applies the table defaults
    setDocTableStyle(dvh, table.id, { mode: 'destination', header: { fill: '#E2EFDA' } })
    refreshDocTable(editor, dvh, table.id)
    const header = pmTableToModel(tableOccurrences(editor.state.doc)[0]!.node.toJSON() as PmNode)
      .rows[0]!
    expect(header[0]?.fill).toBe('E2EFDA')
    expect(dvh.pending.map((c) => c.action)).toEqual([
      'Link.Create',
      'Table.Create',
      'Document.InsertTable',
      'Link.Update',
      'Table.Refresh',
      'Link.Update',
      'Table.Refresh',
      'Table.SetStyle',
      'Table.Refresh',
    ])
  })
})

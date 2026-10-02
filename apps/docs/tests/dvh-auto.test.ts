/**
 * P3: automatic links in Docs. A live message from Sheets or a changed file
 * updates fields and tables without touching the undo history; tables edited
 * by hand are skipped; unlinking leaves plain content.
 */
import { Editor } from '@tiptap/core'
import { undoDepth } from '@tiptap/pm/history'
import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import {
  emptyModel,
  serializeModelXml,
  type DvhCollection,
  type DvhModel,
} from '@genoffice/dvh-model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  fieldOccurrences,
  insertSmartField,
  linkSource,
  loadDvhDocs,
  setActiveDvhDocs,
  setLinkUpdateMode,
  type DvhDocsState,
} from '../src/renderer/dvh-smart-data'
import {
  createDocTable,
  insertDocTable,
  tableOccurrences,
  unlinkSource,
} from '../src/renderer/dvh-tables'
import {
  checkLinksOnOpen,
  installAutoLinks,
  linkStatus,
  refreshAutoWatch,
} from '../src/renderer/dvh-auto'

const editors: Editor[] = []
let stop: (() => void) | null = null
let onLive: ((p: { docId: string; path: string | null; modelXml: string }) => void) | null = null
let onFile: ((path: string) => void) | null = null
let watched: string[] = []
let disk: DvhModel | null = null

beforeEach(() => {
  ;(window as unknown as { desktop: unknown }).desktop = {
    onDvhLive: (cb: typeof onLive) => {
      onLive = cb
      return () => (onLive = null)
    },
    onDvhSourceChanged: (cb: typeof onFile) => {
      onFile = cb
      return () => (onFile = null)
    },
  }
})
afterEach(() => {
  stop?.()
  stop = null
  setActiveDvhDocs(null)
  for (const e of editors.splice(0)) e.destroy()
})

const workItems = (rows: (string | number)[][]): DvhCollection => ({
  id: 'c_work',
  name: 'WorkItems',
  columns: [
    { id: 'k_code', key: 'code', title: 'Mã', type: 'text' },
    { id: 'k_qty', key: 'qty', title: 'KL', type: 'number' },
  ],
  rows,
})
const workbook = (name: string, rows: (string | number)[][]): DvhModel => ({
  ...emptyModel('doc_workbook'),
  fields: [{ id: 'f_name', name: 'Project.Name', type: 'text', value: name, access: 'readwrite' }],
  collections: [workItems(rows)],
})

async function setup(): Promise<{ editor: Editor; dvh: DvhDocsState }> {
  const zip = new JSZip()
  const parsed = {
    internal: { originalBytes: await zip.generateAsync({ type: 'uint8array' }) },
  } as unknown as ParsedDocFull
  const dvh = await loadDvhDocs(parsed)
  setActiveDvhDocs(dvh)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Dự án: ' }] }],
    } as never,
  })
  editors.push(editor)
  disk = workbook('A', [['AB.1', 1]])
  const link = linkSource(dvh, { path: 'D:\\w\\kl.xlsx', model: disk }, 'D:\\w\\bc.docx')
  editor.commands.setTextSelection(editor.state.doc.content.size - 1)
  insertSmartField(editor, dvh, dvh.model!.fields[0]!)
  insertDocTable(editor, dvh, createDocTable(dvh, { collectionId: 'c_work' }), {
    afterBlockIndex: 0,
  })
  setLinkUpdateMode(dvh, link.id, 'auto')
  stop = installAutoLinks({
    editor: () => editor,
    dvh: () => dvh,
    filePath: () => 'D:\\w\\bc.docx',
    readSource: async (path) => ({ path, modelXml: disk ? serializeModelXml(disk) : null }),
    watch: (paths) => (watched = paths),
  })
  refreshAutoWatch()
  return { editor, dvh }
}

const settle = () => new Promise((r) => setTimeout(r, 0))
const rowsOf = (editor: Editor) => {
  const out: string[][] = []
  tableOccurrences(editor.state.doc)[0]!.node.forEach((row) => {
    const cells: string[] = []
    row.forEach((cell) => cells.push(cell.textContent))
    out.push(cells)
  })
  return out
}

describe('automatic links', () => {
  it('watches the workbook and follows live edits without growing the undo history', async () => {
    const { editor, dvh } = await setup()
    expect(watched).toEqual(['D:\\w\\kl.xlsx'])
    const depth = undoDepth(editor.state)
    expect(depth).toBeGreaterThan(0)

    onLive!({
      docId: 'doc_workbook',
      path: 'D:\\w\\kl.xlsx',
      modelXml: serializeModelXml(
        workbook('B', [
          ['AB.1', 1],
          ['AB.2', 2],
        ]),
      ),
    })
    await settle()
    await settle()
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('B')
    expect(rowsOf(editor)).toEqual([
      ['Mã', 'KL'],
      ['AB.1', '1'],
      ['AB.2', '2'],
    ])
    expect(undoDepth(editor.state)).toBe(depth)
    expect(linkStatus(dvh.model!.links[0]!.id)).toBe('current')

    // a workbook the document does not link is ignored
    onLive!({ docId: 'doc_other', path: null, modelXml: serializeModelXml(workbook('X', [])) })
    await settle()
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('B')
  })

  it('a saved file updates the document; a hand-edited table is skipped and flagged', async () => {
    const { editor, dvh } = await setup()
    // type into a body cell
    const occ = tableOccurrences(editor.state.doc)[0]!
    let cellPos = -1
    occ.node.descendants((node, pos) => {
      if (cellPos < 0 && node.isText && node.text === 'AB.1') cellPos = occ.pos + 1 + pos
    })
    editor.view.dispatch(editor.state.tr.insertText('!', cellPos))

    disk = workbook('C', [['AB.9', 9]])
    onFile!('d:\\W\\KL.xlsx')
    await settle()
    await settle()
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('C')
    expect(rowsOf(editor)[1]).toEqual(['!AB.1', '1'])
    expect(linkStatus(dvh.model!.links[0]!.id)).toBe('edited')
  })

  it('on open pulls automatic links and flags stale manual ones; unlink leaves plain content', async () => {
    const { editor, dvh } = await setup()
    const link = dvh.model!.links[0]!
    setLinkUpdateMode(dvh, link.id, 'manual')
    disk = workbook('D', [['AB.1', 1]])
    await checkLinksOnOpen()
    expect(linkStatus(link.id)).toBe('stale')
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('A')
    setLinkUpdateMode(dvh, link.id, 'onOpen')
    await checkLinksOnOpen()
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('D')

    unlinkSource(editor, dvh, link.id)
    expect(fieldOccurrences(editor.state.doc)).toEqual([])
    expect(tableOccurrences(editor.state.doc)).toEqual([])
    expect(editor.state.doc.textContent).toContain('Dự án: D')
    expect(dvh.model!.links).toEqual([])
    expect(dvh.model!.tables).toEqual([])
    expect(dvh.pending.map((c) => c.action)).toContain('Link.Remove')
  })
})

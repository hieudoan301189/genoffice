/**
 * P1: Smart Data in Docs. Fields inserted as bound content controls, field
 * text flowing back into the model (two-way), and a workbook link updating
 * every occurrence.
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
  parseModelXml,
  serializeModelXml,
  type DvhModel,
} from '@genoffice/dvh-model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  dvhDocsCustomXmlParts,
  fieldOccurrences,
  insertSmartField,
  linkSource,
  linkSourcePaths,
  loadDvhDocs,
  setActiveDvhDocs,
  updateFromSource,
} from '../src/renderer/dvh-smart-data'
import { executeOps } from '../src/renderer/ai/ops'
import { createDocsDvhActions } from '../src/renderer/dvh-actions'
import { defaultPermissions } from '@genoffice/dvh-actions'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function editorWith(text: string): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
    } as never,
  })
  editors.push(editor)
  return editor
}

async function parsedWith(parts: Record<string, string> = {}): Promise<ParsedDocFull> {
  const zip = new JSZip()
  for (const [path, xml] of Object.entries(parts)) zip.file(path, xml)
  const originalBytes = await zip.generateAsync({ type: 'uint8array' })
  return { internal: { originalBytes } } as unknown as ParsedDocFull
}

const workbookModel = (value: string): DvhModel => ({
  ...emptyModel('doc_aaaaaaaaaaaaaaaa'),
  fields: [
    {
      id: 'f_aaaaaaaaaaaaaaaa',
      name: 'Project.Name',
      type: 'text',
      value,
      access: 'readwrite',
    },
  ],
})

describe('Docs Smart Data', () => {
  it('reads an existing model part and its store item id', async () => {
    const model = workbookModel('Dự án A')
    const parsed = await parsedWith({
      'customXml/item1.xml': serializeModelXml(model),
      'customXml/itemProps1.xml':
        '<?xml version="1.0"?><ds:datastoreItem ds:itemID="{11111111-2222-3333-4444-555555555555}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"/>',
    })
    const dvh = await loadDvhDocs(parsed)
    expect(dvh.model?.fields[0]?.value).toBe('Dự án A')
    expect(dvh.modelStoreItemId).toBe('{11111111-2222-3333-4444-555555555555}')
    expect(await loadDvhDocs(parsed)).toBe(dvh)
  })

  it('inserts a bound field and writes model and history parts', async () => {
    const parsed = await parsedWith()
    const dvh = await loadDvhDocs(parsed)
    const editor = editorWith('Tên dự án: ')
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    insertSmartField(editor, dvh, workbookModel('Dự án A').fields[0]!)

    const occurrences = fieldOccurrences(editor.state.doc)
    expect(occurrences).toHaveLength(1)
    expect(occurrences[0]).toMatchObject({ fieldId: 'f_aaaaaaaaaaaaaaaa', text: 'Dự án A' })
    expect(occurrences[0]!.sdtPr).toContain(`w:storeItemID="${dvh.modelStoreItemId}"`)

    const parts = dvhDocsCustomXmlParts(parsed, editor)!
    expect(parts.map((p) => p.ns)).toEqual([MODEL_NS, HISTORY_NS])
    expect(parseModelXml(parts[0]!.xml).fields[0]?.value).toBe('Dự án A')
    expect(parseHistoryXml(parts[1]!.xml).changes.map((c) => c.action)).toEqual([
      'Document.InsertField',
    ])
  })

  it('field text edited in the page becomes the model value (two-way)', async () => {
    const parsed = await parsedWith()
    const dvh = await loadDvhDocs(parsed)
    const editor = editorWith('X')
    insertSmartField(editor, dvh, workbookModel('Dự án A').fields[0]!)
    const occ = fieldOccurrences(editor.state.doc)[0]!
    editor.chain().setTextSelection(occ.to).insertContent(' mới').run()
    // non-inclusive mark: text typed at the end stays outside the control
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('Dự án A')
    editor
      .chain()
      .setTextSelection(occ.from + 2)
      .insertContent('Z')
      .run()
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('DựZ án A')

    const parts = dvhDocsCustomXmlParts(parsed, editor)!
    expect(parseModelXml(parts[0]!.xml).fields[0]?.value).toBe('DựZ án A')
    const actions = parseHistoryXml(parts[1]!.xml).changes.map((c) => c.action)
    expect(actions).toEqual(['Document.InsertField', 'Document.EditField'])
    // building the bytes again (a recovery copy, then the save) records nothing new
    const again = dvhDocsCustomXmlParts(parsed, editor)!
    expect(parseHistoryXml(again[1]!.xml).changes).toHaveLength(2)
  })

  it('a link pulls workbook values into every occurrence', async () => {
    const parsed = await parsedWith()
    const dvh = await loadDvhDocs(parsed)
    const editor = editorWith('A')
    const link = linkSource(
      dvh,
      { path: 'D:\\Work\\boq.xlsx', model: workbookModel('Dự án A') },
      'D:\\Work\\report.docx',
    )
    expect(link.source).toMatchObject({ relPath: 'boq.xlsx', path: 'D:\\Work\\boq.xlsx' })
    const field = dvh.model!.fields[0]!
    insertSmartField(editor, dvh, field)
    editor.commands.setTextSelection(1)
    insertSmartField(editor, dvh, field)
    expect(fieldOccurrences(editor.state.doc).map((o) => o.text)).toEqual(['Dự án A', 'Dự án A'])

    const changed = updateFromSource(editor, dvh, link, {
      path: 'D:\\Work\\boq.xlsx',
      model: workbookModel('Dự án B'),
    })
    expect(changed).toEqual({ fields: 2, collections: [], localEdits: [], conflicts: [] })
    expect(fieldOccurrences(editor.state.doc).map((o) => o.text)).toEqual(['Dự án B', 'Dự án B'])
    expect(link.lastSync?.revision).toBe(1)
    const parts = dvhDocsCustomXmlParts(parsed, editor)!
    const saved = parseModelXml(parts[0]!.xml)
    expect(saved.fields[0]?.value).toBe('Dự án B')
    expect(saved.links[0]?.source.docId).toBe('doc_aaaaaaaaaaaaaaaa')
    expect(parseHistoryXml(parts[1]!.xml).changes.map((c) => c.action)).toContain('Link.Update')
  })

  it('looks for a linked workbook next to the document first', () => {
    const link = {
      id: 'l_aaaaaaaaaaaaaaaa',
      source: {
        docId: 'doc_aaaaaaaaaaaaaaaa',
        relPath: 'boq.xlsx',
        path: 'D:\\Old\\boq.xlsx',
        objectId: '*',
      },
      targets: [],
      update: 'manual' as const,
    }
    expect(linkSourcePaths(link, 'E:\\Moved\\report.docx')).toEqual([
      'E:\\Moved\\boq.xlsx',
      'D:\\Old\\boq.xlsx',
    ])
    expect(linkSourcePaths(link, null)).toEqual(['D:\\Old\\boq.xlsx'])
  })
})

describe('Smart Data ops and actions', () => {
  async function activeDoc(text: string) {
    const parsed = await parsedWith()
    const dvh = await loadDvhDocs(parsed)
    setActiveDvhDocs(dvh)
    return { parsed, dvh, editor: editorWith(text) }
  }
  afterEach(() => setActiveDvhDocs(null))

  it('insertSmartField and setSmartField run through apply_ops as one batch', async () => {
    const { parsed, dvh, editor } = await activeDoc('Dự án: ')
    const outcome = executeOps(editor, [
      { op: 'insertSmartField', target: { blockIndexes: [0] }, field: 'Project.Name', value: 'A' },
      {
        op: 'insertSmartField',
        target: { blockIndexes: [0] },
        field: 'Project.Name',
        position: 'start',
      },
    ])
    expect(outcome.ok).toBe(true)
    expect(fieldOccurrences(editor.state.doc).map((o) => o.text)).toEqual(['A', 'A'])
    expect(dvh.model?.fields.map((f) => f.name)).toEqual(['Project.Name'])

    expect(
      executeOps(editor, [{ op: 'setSmartField', field: 'Project.Name', value: 'B' }]).ok,
    ).toBe(true)
    expect(fieldOccurrences(editor.state.doc).map((o) => o.text)).toEqual(['B', 'B'])
    expect(dvh.model?.fields[0]?.value).toBe('B')
    const actions = parseHistoryXml(dvhDocsCustomXmlParts(parsed, editor)![1]!.xml).changes.map(
      (c) => c.action,
    )
    expect(actions).toEqual(['Document.InsertField', 'Document.InsertField', 'Data.SetField'])
  })

  it('a rejected batch leaves the model untouched', async () => {
    const { dvh, editor } = await activeDoc('X')
    const outcome = executeOps(editor, [
      { op: 'insertSmartField', target: { blockIndexes: [0] }, field: 'Project.Name', value: 'A' },
      { op: 'insertSmartField', target: { blockIndexes: [7] }, field: 'Project.Name', value: 'A' },
    ])
    expect(outcome.ok).toBe(false)
    expect(dvh.model?.fields ?? []).toEqual([])
    expect(fieldOccurrences(editor.state.doc)).toEqual([])
    const unknown = executeOps(editor, [{ op: 'setSmartField', field: 'Nope', value: 1 }])
    expect(unknown.error).toContain('unknown field "Nope"')
  })

  it('the action registry lists, sets and inserts fields', async () => {
    const { editor } = await activeDoc('Hạng mục ')
    const registry = createDocsDvhActions({
      editor: () => editor,
      filePath: () => null,
      readSource: async () => null,
    })
    const ai = { caller: 'ai' as const, docId: 'doc_x', permissions: defaultPermissions('ai') }
    const inserted = await registry.run(
      'Document.InsertField',
      { field: 'Item.Name', blockIndex: 0, value: 'Móng' },
      ai,
    )
    expect(inserted.changeSet?.action).toBe('Document.InsertField')
    const set = await registry.run('Data.SetField', { field: 'Item.Name', value: 'Thân' }, ai)
    expect(set.changeSet?.changes[0]).toMatchObject({ before: 'Móng', after: 'Thân' })
    const list = await registry.run<{ name: string; value: unknown; uses: number }[]>(
      'Data.ListFields',
      {},
      ai,
    )
    expect(list.output).toEqual([
      expect.objectContaining({ name: 'Item.Name', value: 'Thân', uses: 1 }),
    ])
    expect(
      registry
        .catalog()
        .map((a) => a.name)
        .slice(0, 8),
    ).toEqual([
      'Data.ListFields',
      'Data.GetField',
      'Data.SetField',
      'Document.InsertField',
      'Link.Update',
      'Document.InsertTable',
      'Table.Refresh',
      'Table.SetStyle',
    ])
    // P5 history actions, then (P4) the editor's op system as Document.* actions
    const names = registry.catalog().map((a) => a.name)
    expect(names.slice(8, 11)).toEqual([
      'History.RestoreObject',
      'History.RevertTransaction',
      'Table.ReplaceData',
    ])
    expect(names.slice(11).every((name) => name.startsWith('Document.'))).toBe(true)
  })
})

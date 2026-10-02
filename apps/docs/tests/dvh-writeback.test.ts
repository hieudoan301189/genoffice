/**
 * P3 acceptance, Docs side: a Read/Write field edited in the document goes
 * back to the workbook (round-trip through real xlsx bytes, with a change set
 * on both sides); a field changed on both sides is a conflict the user
 * settles; automatic links write back after typing; a renamed or moved
 * workbook is found again by its docId, and two candidates are reported.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { readCustomXmlPart } from '@genoffice/docx-engine'
import {
  HISTORY_NS,
  MODEL_NS,
  parseHistoryXml,
  parseModelXml,
  serializeModelXml,
  type DvhModel,
  type WriteFieldsRequest,
  type WriteFieldsResult,
} from '@genoffice/dvh-model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  fieldOccurrences,
  insertSmartField,
  linkSource,
  loadDvhDocs,
  pendingWriteBacks,
  setActiveDvhDocs,
  setLinkUpdateMode,
  type DvhDocsState,
} from '../src/renderer/dvh-smart-data'
import {
  ambiguousSources,
  checkLinksOnOpen,
  installAutoLinks,
  linkConflicts,
  linkStatus,
  noteDocumentEdited,
  resetLinkState,
  resolveConflict,
  writeBackLink,
} from '../src/renderer/dvh-auto'
import { modelFromCells } from '../src/main/dvh-xlsx-cells'
import { writeFieldsToXlsx } from '../src/main/dvh-xlsx-write'

const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const WB = 'doc_workbook00000001'

const sourceModel: DvhModel = {
  docId: WB,
  schemaVersion: 1,
  fields: [
    { id: 'f_name', name: 'Project.Name', type: 'text', value: 'Dự án A', access: 'readwrite' },
    { id: 'f_qty', name: 'Lot.Qty', type: 'number', value: 5, access: 'readwrite' },
  ],
  collections: [],
  tables: [],
  links: [],
}

async function workbookBytes(name: string, qty: number): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`,
  )
  zip.file(
    'xl/workbook.xml',
    `${X}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Info" sheetId="1" r:id="rId1"/></sheets>` +
      '<definedNames><definedName name="_dvh.f.f_name" hidden="1">Info!$B$1</definedName>' +
      '<definedName name="_dvh.f.f_qty" hidden="1">Info!$B$2</definedName></definedNames></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${X}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` +
      `<row r="1"><c r="B1" t="inlineStr"><is><t>${name}</t></is></c></row>` +
      `<row r="2"><c r="B2"><v>${qty}</v></c></row></sheetData></worksheet>`,
  )
  zip.file('customXml/item1.xml', serializeModelXml(sourceModel))
  return zip.generateAsync({ type: 'uint8array' })
}

/** a tiny disk: path → workbook bytes, read as the main process does (cells are the truth) */
let disk = new Map<string, Uint8Array>()
async function readSource(path: string) {
  const bytes = disk.get(path)
  if (!bytes) throw new Error('ENOENT')
  const part = await readCustomXmlPart(bytes, MODEL_NS)
  const model = await modelFromCells(bytes, parseModelXml(part!.xml))
  return { path, modelXml: serializeModelXml(model) }
}
/** the main process' file path of a write (no Sheets tab open) */
async function writeFields(request: WriteFieldsRequest): Promise<WriteFieldsResult> {
  for (const path of request.paths) {
    const bytes = disk.get(path)
    if (!bytes) continue
    const out = await writeFieldsToXlsx(bytes, request)
    if (out.bytes) disk.set(path, out.bytes)
    return { via: 'file', path, results: out.results }
  }
  return { via: 'none', path: null, results: [], error: 'not found' }
}
let index: string[] = []

const editors: Editor[] = []
let stop: (() => void) | null = null
beforeEach(() => {
  ;(window as unknown as { desktop: unknown }).desktop = {
    onDvhLive: () => () => {},
    onDvhSourceChanged: () => () => {},
  }
  disk = new Map()
  index = []
  resetLinkState()
})
afterEach(() => {
  stop?.()
  stop = null
  setActiveDvhDocs(null)
  for (const e of editors.splice(0)) e.destroy()
})

const SRC = 'D:\\w\\nguon.xlsx'
const DOCX = 'D:\\w\\bao-cao.docx'

async function setup(): Promise<{ editor: Editor; dvh: DvhDocsState; linkId: string }> {
  disk.set(SRC, await workbookBytes('Dự án A', 5))
  const parsed = {
    internal: { originalBytes: await new JSZip().generateAsync({ type: 'uint8array' }) },
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
  const source = (await readSource(SRC))!
  const link = linkSource(dvh, { path: SRC, model: parseModelXml(source.modelXml) }, DOCX)
  editor.commands.setTextSelection(editor.state.doc.content.size - 1)
  insertSmartField(editor, dvh, dvh.model!.fields[0]!)
  stop = installAutoLinks({
    editor: () => editor,
    dvh: () => dvh,
    filePath: () => DOCX,
    readSource,
    watch: () => {},
    writeFields,
    findSource: async (docId) => {
      const found: string[] = []
      for (const path of index) {
        const bytes = disk.get(path)
        const part = bytes ? await readCustomXmlPart(bytes, MODEL_NS) : null
        if (part && parseModelXml(part.xml).docId === docId) found.push(path)
      }
      return found
    },
  })
  return { editor, dvh, linkId: link.id }
}

/** types over the first Smart Field occurrence, as a user would */
function typeField(editor: Editor, text: string): void {
  const occ = fieldOccurrences(editor.state.doc)[0]!
  const marks = editor.state.doc.nodeAt(occ.from)!.marks
  editor.view.dispatch(
    editor.state.tr.replaceWith(occ.from, occ.to, editor.schema.text(text, marks)),
  )
}

async function sourceState(path = SRC) {
  const bytes = disk.get(path)!
  const model = await modelFromCells(
    bytes,
    parseModelXml((await readCustomXmlPart(bytes, MODEL_NS))!.xml),
  )
  const historyPart = await readCustomXmlPart(bytes, HISTORY_NS)
  return {
    model,
    rev: parseModelXml((await readCustomXmlPart(bytes, MODEL_NS))!.xml).fields[0]!.rev,
    history: historyPart ? parseHistoryXml(historyPart.xml).changes : [],
  }
}

describe('field write-back (Read/Write links)', () => {
  it('round-trip: an edit in the document reaches the workbook, with history on both sides', async () => {
    const { editor, dvh, linkId } = await setup()
    const link = dvh.model!.links[0]!
    typeField(editor, 'Dự án B')
    expect(pendingWriteBacks(editor.state.doc, dvh, link)).toEqual(['f_name'])

    const outcome = await writeBackLink(linkId)
    expect(outcome).toEqual({ written: 1, conflicts: 0, refused: [] })
    const src = await sourceState()
    expect(src.model.fields[0]!.value).toBe('Dự án B')
    expect(src.rev).toBe(1)
    expect(src.history.map((c) => c.action)).toEqual(['Data.SetField'])
    expect(dvh.pending.map((c) => c.action)).toContain('Link.WriteBack')
    expect(pendingWriteBacks(editor.state.doc, dvh, link)).toEqual([])
    expect(link.lastSync!.fields!.f_name).toEqual({ rev: 1, text: 'Dự án B' })

    // the next pull agrees: nothing to change, no conflict
    await checkLinksOnOpen()
    expect(linkStatus(linkId)).toBe('current')
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('Dự án B')
  })

  it('a field changed on both sides is a conflict: keep source, or keep this side', async () => {
    const { editor, dvh, linkId } = await setup()
    const link = dvh.model!.links[0]!
    setLinkUpdateMode(dvh, linkId, 'onOpen')
    typeField(editor, 'Sửa ở Docs')
    // meanwhile Excel changes the cell and saves
    disk.set(SRC, await workbookBytes('Sửa ở Excel', 5))

    // the pull keeps the document's text and reports the conflict
    await checkLinksOnOpen()
    expect(linkStatus(linkId)).toBe('conflict')
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('Sửa ở Docs')
    expect(linkConflicts(linkId)).toEqual([
      {
        fieldId: 'f_name',
        local: 'Sửa ở Docs',
        source: { rev: 0, text: 'Sửa ở Excel' },
        base: { rev: 0, text: 'Dự án A' },
      },
    ])

    // a plain write-back refuses to overwrite it too
    expect((await writeBackLink(linkId))!.written).toBe(0)
    expect((await sourceState()).model.fields[0]!.value).toBe('Sửa ở Excel')

    // keep source
    await resolveConflict(linkId, 'f_name', 'source')
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('Sửa ở Excel')
    expect(linkStatus(linkId)).toBe('current')
    expect(linkConflicts(linkId)).toEqual([])
    expect(dvh.pending.map((c) => c.action)).toContain('Link.ResolveConflict')

    // again, this time keep this side: forced over the source
    typeField(editor, 'Docs thắng')
    disk.set(SRC, await workbookBytes('Excel lần 2', 5))
    await checkLinksOnOpen()
    expect(linkStatus(linkId)).toBe('conflict')
    await resolveConflict(linkId, 'f_name', 'mine')
    expect((await sourceState()).model.fields[0]!.value).toBe('Docs thắng')
    expect(linkStatus(linkId)).toBe('current')
    expect(link.lastSync!.fields!.f_name!.text).toBe('Docs thắng')
  })

  it('automatic links write back shortly after typing stops', async () => {
    const { editor, dvh, linkId } = await setup()
    setLinkUpdateMode(dvh, linkId, 'auto')
    typeField(editor, 'Gõ tự động')
    noteDocumentEdited()
    await new Promise((r) => setTimeout(r, 1500))
    expect((await sourceState()).model.fields[0]!.value).toBe('Gõ tự động')
  })

  it('relinks a renamed workbook by docId; two copies are reported, not guessed', async () => {
    const { editor, dvh, linkId } = await setup()
    const link = dvh.model!.links[0]!
    const moved = 'E:\\du-an\\nguon-2026.xlsx'
    disk.set(moved, disk.get(SRC)!)
    disk.delete(SRC)
    index = [moved]
    await checkLinksOnOpen()
    expect(linkStatus(linkId)).toBe('current')
    expect(link.source.path).toBe(moved)
    expect(link.source.relPath).toBeUndefined()
    expect(dvh.pending.map((c) => c.action)).toContain('Link.Relocate')

    // write-back follows the new place
    typeField(editor, 'Sau khi chuyển')
    await writeBackLink(linkId)
    expect((await sourceState(moved)).model.fields[0]!.value).toBe('Sau khi chuyển')

    // gone again, and now two copies carry the docId
    const a = 'F:\\a\\nguon.xlsx'
    const b = 'F:\\b\\nguon.xlsx'
    disk.set(a, disk.get(moved)!)
    disk.set(b, disk.get(moved)!)
    disk.delete(moved)
    index = [a, b]
    await checkLinksOnOpen()
    expect(linkStatus(linkId)).toBe('ambiguous')
    expect(ambiguousSources(linkId)).toEqual([a, b])
    expect(link.source.path).toBe(moved)

    // nothing anywhere: missing
    disk.clear()
    await checkLinksOnOpen()
    expect(linkStatus(linkId)).toBe('missing')
  })
})

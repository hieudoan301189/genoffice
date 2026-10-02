/**
 * P10 in Docs: the project data service of the main process (schema pack,
 * objects, the `dvh-project://` source, field write-back with conflicts, the
 * two-way sync with a QLCL workbook on disk) and a document linked straight
 * to project objects, following their changes.
 */
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { parseModelXml } from '@genoffice/dvh-model'
import { projectFieldId } from '@genoffice/dvh-project'
import { exportQlclWorkbook, importQlclWorkbook } from '@genoffice/dvh-template'
import { ProjectDataService, type ProjectDataStore } from '../src/main/dvh-project-source'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  fieldOccurrences,
  insertSmartField,
  linkSource,
  loadDvhDocs,
  setActiveDvhDocs,
  sourceFromRead,
} from '../src/renderer/dvh-smart-data'
import { applySource } from '../src/renderer/dvh-auto'

const PACK = readFileSync(
  join(__dirname, '../../../packages/dvh-project/packs/qlcl-xay-dung/pack.json'),
  'utf8',
)

function memoryStore(): ProjectDataStore & { files: Map<string, unknown> } {
  const files = new Map<string, unknown>()
  return {
    files,
    resolveProjectForFile: () => 'p1',
    ensureDefaultProject: () => ({ id: 'default', name: 'Default Project' }),
    getProject: (id) => (id === 'p1' ? { id, name: 'Cầu Rạch Miễu 2' } : null),
    readProjectData: (id) => structuredClone(files.get(id) ?? null),
    writeProjectData: (id, data) => {
      files.set(id, structuredClone(data))
    },
  }
}

const editors: Editor[] = []
const dirs: string[] = []
afterEach(() => {
  setActiveDvhDocs(null)
  for (const e of editors.splice(0)) e.destroy()
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true })
})

function service() {
  const changed: string[] = []
  const store = memoryStore()
  const svc = new ProjectDataService(
    () => store,
    (uri) => changed.push(uri),
    () => '2026-10-02T10:00:00.000Z',
  )
  svc.installPack('p1', PACK)
  const projectId = svc.setObject('p1', {
    type: 'Project',
    objectId: null,
    values: { Code: 'DA-01', Name: 'Cầu Rạch Miễu 2' },
  })
  const item = svc.setObject('p1', {
    type: 'WorkItem',
    objectId: null,
    values: { Code: 'HM-01', Name: 'Móng trụ T1', Quantity: '125' },
  })
  return { svc, store, changed, projectId, item }
}

describe('project data in Docs', () => {
  it('serves packs, objects, the project source and checked write-back', () => {
    const { svc, changed, projectId } = service()
    const info = svc.info('p1')
    expect(info).toMatchObject({
      name: 'Cầu Rạch Miễu 2',
      uri: 'dvh-project://p1',
      docId: 'proj_p1',
    })
    expect(info.packs).toEqual([
      { id: 'qlcl-xay-dung', name: 'QLCL công trình xây dựng', packVersion: '1.0.0' },
    ])
    expect(info.types.find((t) => t.name === 'WorkItem')).toMatchObject({ single: false, count: 1 })
    expect(changed.every((u) => u === 'dvh-project://p1')).toBe(true)
    expect(svc.objects('p1', 'WorkItem')[0]!.values).toMatchObject({ Quantity: 125 })

    const read = svc.readSource('dvh-project://p1')!
    const model = parseModelXml(read.modelXml!)
    expect(model.fields.find((f) => f.name === 'Project.Name')!.value).toBe('Cầu Rạch Miễu 2')
    expect(svc.readSource('dvh-project://other')).toBeNull()

    const fieldId = projectFieldId(projectId, 'Name')
    const request = (value: string, rev: number) => ({
      docId: 'proj_p1',
      paths: ['dvh-project://p1'],
      writes: [{ fieldId, value, expected: { rev, text: 'Cầu Rạch Miễu 2' }, force: false }],
      origin: { docId: 'doc_x', name: 'BB-01.docx' },
    })
    expect(svc.writeFields(request('Cầu RM II', 1))!.results[0]).toMatchObject({
      status: 'written',
    })
    expect(svc.writeFields(request('Cầu khác', 1))!.results[0]).toMatchObject({
      status: 'conflict',
    })
    // not a project path: the workbook writer handles it
    expect(svc.writeFields({ ...request('x', 1), paths: ['D:\\a.xlsx'] })).toBeNull()
  })

  it('syncs both ways with a QLCL workbook on disk; conflicts wait for the user', async () => {
    const { svc, item } = service()
    const dir = mkdtempSync(join(tmpdir(), 'dvh-project-'))
    dirs.push(dir)
    const partner = join(dir, 'QLCL-DA01.xlsx')
    // QLCL-DVH starts from an export of ours and edits a quantity; we edit the same one differently
    const first = await svc.sync(
      'p1',
      (writeFileSync(
        partner,
        await exportQlclWorkbook({
          docId: 'doc_q',
          schemaVersion: 1,
          fields: [],
          collections: [],
          tables: [],
          links: [],
        }),
      ),
      partner),
    )
    expect(first.conflicts).toEqual([])
    const theirs = await importQlclWorkbook(new Uint8Array(readFileSync(partner)))
    const items = theirs.collections.find((c) => c.name === 'WorkItem')!
    const col = items.columns.findIndex((c) => c.title === 'Khối lượng')
    items.rows[0]![col] = 130
    writeFileSync(partner, await exportQlclWorkbook(theirs))
    svc.setObject('p1', { type: 'WorkItem', objectId: item, values: { Quantity: 128 } })
    const second = await svc.sync('p1', partner)
    expect(second.conflicts).toEqual([
      expect.objectContaining({ objectId: item, key: 'Quantity', local: '128', remote: '130' }),
    ])
    expect(svc.info('p1').partners[0]!.conflicts).toHaveLength(1)
    // the partner file keeps its own value
    const written = await importQlclWorkbook(new Uint8Array(readFileSync(partner)))
    expect(written.collections.find((c) => c.name === 'WorkItem')!.rows[0]![col]).toBe(130)
    svc.resolve('p1', partner, 0, 'mine')
    expect(svc.info('p1').partners[0]!.conflicts).toEqual([])
    await svc.sync('p1', partner)
    const pushed = await importQlclWorkbook(new Uint8Array(readFileSync(partner)))
    expect(pushed.collections.find((c) => c.name === 'WorkItem')!.rows[0]![col]).toBe(128)
  })

  it('a document links straight to project objects and follows them', async () => {
    const { svc, projectId } = service()
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
    const link = linkSource(
      dvh,
      sourceFromRead(svc.readSource('dvh-project://p1'))!,
      'D:\\HoSo\\BB-01.docx',
    )
    expect(link.source).toMatchObject({ docId: 'proj_p1', path: 'dvh-project://p1' })
    expect(link.source.relPath).toBeUndefined()
    editor.commands.setTextSelection(8)
    insertSmartField(
      editor,
      dvh,
      dvh.model!.fields.find((f) => f.name === 'Project.Name')!,
    )
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('Cầu Rạch Miễu 2')
    // the project changes (another document, the panel, a sync): the link pulls it
    svc.setObject('p1', {
      type: 'Project',
      objectId: projectId,
      values: { Name: 'Cầu Rạch Miễu II' },
    })
    applySource(editor, dvh, link, sourceFromRead(svc.readSource('dvh-project://p1'))!, false)
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('Cầu Rạch Miễu II')
    expect(dvh.model!.collections.find((c) => c.name === 'WorkItem')!.rows[0]![1]).toBe('HM-01')
  })
})

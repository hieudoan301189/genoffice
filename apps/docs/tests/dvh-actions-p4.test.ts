/**
 * P4 in Docs: the editor's op system is in the catalog as Document.* actions
 * (preview = dry run, execute = one undo item), and a saga that fails in a
 * second document puts this one back to its checkpoint.
 */
import { Editor } from '@tiptap/core'
import { undoDepth } from '@tiptap/pm/history'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { defaultPermissions, runSaga, type SagaParticipant } from '@genoffice/dvh-actions'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { loadDvhDocs, setActiveDvhDocs } from '../src/renderer/dvh-smart-data'
import { createDocsDvhActions, docsSagaParticipant } from '../src/renderer/dvh-actions'

const editors: Editor[] = []
afterEach(() => {
  setActiveDvhDocs(null)
  for (const e of editors.splice(0)) e.destroy()
})

async function openDoc(texts: string[]) {
  const parsed = {
    internal: { originalBytes: await new JSZip().generateAsync({ type: 'uint8array' }) },
  } as unknown as ParsedDocFull
  setActiveDvhDocs(await loadDvhDocs(parsed))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: texts.map((text) => ({ type: 'docParagraph', content: [{ type: 'text', text }] })),
    } as never,
  })
  editors.push(editor)
  const host = { editor: () => editor, filePath: () => null, readSource: async () => null }
  return { editor, host, registry: createDocsDvhActions(host) }
}

const ui = {
  caller: 'ui' as const,
  docId: 'doc_a',
  permissions: defaultPermissions('ui'),
  confirm: () => true,
}

describe('Document.* adapter', () => {
  it('wraps OpDef ops: dry-run preview, one undo item per execute', async () => {
    const { editor, registry } = await openDoc(['Báo cáo tuần', 'Nội dung'])
    expect(registry.has('Document.SetFont')).toBe(true)
    expect(registry.get('Document.DeleteBlocks')!.effect).toBe('destructive')
    const target = { blockIndexes: [0] }
    const dry = await registry.run(
      'Document.SetFont',
      { target, bold: true },
      { ...ui, dryRun: true },
    )
    expect(dry.preview.summary[0]).toContain('setFont')
    expect(editor.getHTML()).not.toContain('<strong>')
    const depth = undoDepth(editor.state)
    const done = await registry.run('Document.SetFont', { target, bold: true }, ui)
    expect(editor.getHTML()).toContain('<strong>Báo cáo tuần</strong>')
    expect(undoDepth(editor.state)).toBe(depth + 1)
    expect(done.changeSet?.action).toBe('Document.SetFont')
    // the op system's own validation still applies
    await expect(registry.run('Document.SetFont', { bold: true }, ui)).rejects.toThrow(/target/)
  })

  it('a saga failing in the second document restores the first', async () => {
    const a = await openDoc(['Một', 'Hai'])
    const before = a.editor.getHTML()
    const broken: SagaParticipant = {
      docId: 'doc_b',
      run: async () => {
        throw new Error('workbook locked')
      },
      checkpoint: async () => null,
      restore: async () => {},
    }
    const participants = new Map<string, SagaParticipant>([
      ['doc_a', docsSagaParticipant(a.host, a.registry, 'doc_a')],
      ['doc_b', broken],
    ])
    const result = await runSaga(
      [
        { docId: 'doc_a', action: 'Document.FindReplace', input: { find: 'Một', replace: 'Ba' } },
        { docId: 'doc_b', action: 'Spreadsheet.SetCells', input: {} },
      ],
      participants,
      { caller: 'workflow', permissions: defaultPermissions('ui'), confirm: () => true },
    )
    expect(result.ok).toBe(false)
    expect(result.rolledBack).toEqual(['doc_b', 'doc_a'])
    expect(a.editor.getHTML()).toBe(before)
  })
})

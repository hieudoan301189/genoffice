/**
 * P7 in Docs: an AI plan runs through the catalog as one transaction; the
 * history records its change sets under the run's txId with source `ai`, and
 * History.RevertTransaction undoes the whole AI operation.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { createDvhPlannerSkill } from '@genoffice/dvh-actions'
import { emptyModel } from '@genoffice/dvh-model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  fieldOccurrences,
  insertSmartField,
  linkSource,
  loadDvhDocs,
  setActiveDvhDocs,
} from '../src/renderer/dvh-smart-data'
import { createDocsDvhActions, docsSagaParticipant } from '../src/renderer/dvh-actions'
import { revertTransaction } from '../src/renderer/dvh-history'

const editors: Editor[] = []
afterEach(() => {
  setActiveDvhDocs(null)
  for (const e of editors.splice(0)) e.destroy()
})

describe('AI plans in Docs', () => {
  it('one txId for the whole AI run, undone as a whole', async () => {
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
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'A: ' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'B: ' }] },
        ],
      } as never,
    })
    editors.push(editor)
    linkSource(
      dvh,
      {
        path: 'D:\\w\\nguon.xlsx',
        model: {
          ...emptyModel('doc_workbook0000001'),
          fields: [
            { id: 'f_a', name: 'Project.Name', type: 'text', value: 'A0', access: 'readwrite' },
            { id: 'f_b', name: 'Project.Owner', type: 'text', value: 'B0', access: 'readwrite' },
          ],
        },
      },
      null,
    )
    editor.commands.setTextSelection(4)
    insertSmartField(editor, dvh, dvh.model!.fields[0]!)
    const host = { editor: () => editor, filePath: () => null, readSource: async () => null }
    const registry = createDocsDvhActions(host)
    const participant = docsSagaParticipant(host, registry, dvh.model!.docId)
    const skill = createDvhPlannerSkill({
      registry: () => registry,
      docId: () => dvh.model!.docId,
      model: () => dvh.model,
      checkpoint: () => participant.checkpoint(),
      restore: (cp) => participant.restore(cp),
      confirm: () => true,
    })
    const result = await skill.executeTool({
      id: 'c',
      name: 'propose_plan',
      input: {
        summary: 'đổi tên dự án và chủ đầu tư',
        steps: [
          { action: 'Data.SetField', input: { field: 'Project.Name', value: 'A1' } },
          { action: 'Data.SetField', input: { field: 'Project.Owner', value: 'B1' } },
        ],
      },
    })
    const { txId } = JSON.parse(result.output) as { txId: string }
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('A1')
    const run = dvh.pending.filter((c) => c.txId === txId)
    expect(run.length).toBeGreaterThanOrEqual(2)
    expect(run.every((c) => c.source === 'ai')).toBe(true)

    const undone = revertTransaction(editor, dvh, txId)
    expect(undone.conflicts).toEqual([])
    expect(fieldOccurrences(editor.state.doc)[0]!.text).toBe('A0')
    expect(dvh.model!.fields.find((f) => f.id === 'f_b')!.value).toBe('B0')
  })
})

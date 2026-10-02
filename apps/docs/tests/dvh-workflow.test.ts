/**
 * P8 in Docs: the panel's commands reach the recorder through the UI action
 * registry; a recorded workflow is stored in the document's model and runs
 * again (Workflow.Run) with another record as one transaction — rolled back
 * when a step fails, undone as a whole through History.RevertTransaction.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { emptyModel, parseModelXml, serializeModelXml } from '@genoffice/dvh-model'
import { loadWorkflows } from '@genoffice/dvh-workflow'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  fieldOccurrences,
  insertSmartField,
  linkSource,
  loadDvhDocs,
  setActiveDvhDocs,
} from '../src/renderer/dvh-smart-data'
import { revertTransaction } from '../src/renderer/dvh-history'
import {
  announceUi,
  docsRecorder,
  installDocsUiActions,
  noteUnrecordable,
  resetDocsUiActions,
  runUiAction,
  setRecordSelection,
  startRecording,
  stopRecording,
} from '../src/renderer/dvh-workflow'

const editors: Editor[] = []
afterEach(() => {
  resetDocsUiActions()
  setActiveDvhDocs(null)
  for (const e of editors.splice(0)) e.destroy()
})

async function setup() {
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
  const registry = installDocsUiActions({
    editor: () => editor,
    filePath: () => null,
    readSource: async () => null,
    confirm: () => true,
  })
  const text = () => fieldOccurrences(editor.state.doc)[0]!.text
  const value = (id: string) => dvh.model!.fields.find((f) => f.id === id)!.value
  return { dvh, editor, registry, text, value }
}

describe('workflows in Docs', () => {
  it('records panel commands, stores the workflow and runs it again with another record', async () => {
    const { dvh, editor, registry, text, value } = await setup()
    setRecordSelection({ name: 'Cầu Rạch Miễu', owner: 'Ban QLDA 7' })
    startRecording()
    await runUiAction('Data.SetField', { field: 'f_a', value: 'Cầu' })
    await runUiAction('Data.SetField', { field: 'f_a', value: 'Cầu Rạch Miễu' })
    await runUiAction('Data.ListFields', {})
    noteUnrecordable('Link workbook')
    await runUiAction('Data.SetField', { field: 'f_b', value: 'Ban QLDA 7' })
    // a command with its own code path announces the catalog action it amounts to
    announceUi('Table.Refresh', { table: 'tbl_none' })
    announceUi('Table.Refresh', { bad: true })
    stopRecording()
    expect(docsRecorder.steps().map((s) => s.action)).toEqual([
      'Data.SetField',
      'Data.SetField',
      'Table.Refresh',
    ])
    expect(docsRecorder.unrecorded()).toEqual(['Link workbook'])
    docsRecorder.clear()
    startRecording()
    await runUiAction('Data.SetField', { field: 'f_a', value: 'Cầu Rạch Miễu' })
    await runUiAction('Data.SetField', { field: 'f_b', value: 'Ban QLDA 7' })
    stopRecording()
    expect(text()).toBe('Cầu Rạch Miễu')

    const workflow = docsRecorder.toWorkflow({ name: 'Đặt dự án', id: 'wf_project' })
    await runUiAction('Workflow.Save', { workflow })
    // stored in the model part: survives a save and reopen
    const reopened = parseModelXml(serializeModelXml(dvh.model!))
    expect(loadWorkflows(reopened).workflows).toEqual([workflow])

    // the agent runs it for another record; its steps are not recorded
    startRecording()
    const run = await registry.run(
      'Workflow.Run',
      { workflow: 'Đặt dự án', params: { record: { name: 'Cầu Mỹ Thuận', owner: 'Ban QLDA 9' } } },
      { caller: 'ai', docId: dvh.model!.docId, permissions: new Set(['read', 'write']) },
    )
    stopRecording()
    expect(docsRecorder.steps()).toEqual([])
    expect(text()).toBe('Cầu Mỹ Thuận')
    expect(value('f_b')).toBe('Ban QLDA 9')
    expect(run.output).toMatchObject({ actions: 2 })
    const changeSets = dvh.pending.filter(
      (c) => c.action === 'Data.SetField' && c.source === 'workflow',
    )
    expect(new Set(changeSets.map((c) => c.txId)).size).toBe(1)

    // the whole run is undone as one transaction
    const undone = revertTransaction(editor, dvh, changeSets[0]!.txId)
    expect(undone.conflicts).toEqual([])
    expect(text()).toBe('Cầu Rạch Miễu')
    expect(value('f_b')).toBe('Ban QLDA 7')
  })

  it('a failing step rolls the whole workflow back', async () => {
    const { dvh, registry, text, value } = await setup()
    await runUiAction('Workflow.Save', {
      workflow: {
        format: 'dvh-workflow',
        version: 1,
        id: 'wf_bad',
        name: 'Hỏng',
        params: [],
        steps: [
          {
            kind: 'action',
            action: 'Data.SetField',
            args: { field: { value: 'f_a' }, value: { value: 'X' } },
          },
          {
            kind: 'action',
            action: 'Data.SetField',
            args: { field: { value: 'nope' }, value: { value: 'Y' } },
          },
        ],
      },
    })
    await expect(
      registry.run(
        'Workflow.Run',
        { workflow: 'wf_bad' },
        { caller: 'ui', docId: dvh.model!.docId, permissions: new Set(['read', 'write']) },
      ),
    ).rejects.toThrow(/stopped at step 1: unknown field "nope"/)
    expect(text()).toBe('A0')
    expect(value('f_a')).toBe('A0')
    // unknown actions are refused when saving
    await expect(
      runUiAction('Workflow.Save', {
        workflow: {
          format: 'dvh-workflow',
          version: 1,
          id: 'wf_x',
          name: 'x',
          params: [],
          steps: [{ kind: 'action', action: 'File.DeleteAll', args: {} }],
        },
      }),
    ).rejects.toThrow(/unknown action\(s\): File.DeleteAll/)
  })
})

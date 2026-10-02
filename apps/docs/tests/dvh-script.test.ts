/**
 * P9 in Docs: the agent's script skill writes DVH-Script that is checked
 * against the Docs catalog, saved as a workflow of the document, or run as
 * one transaction in the editor; the script editor's text round-trips through
 * the block editor; a workflow that runs itself stops at the nesting limit.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { emptyModel } from '@genoffice/dvh-model'
import { createDvhScriptSkill, parseScript, printScript } from '@genoffice/dvh-script'
import { loadWorkflows } from '@genoffice/dvh-workflow'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  fieldOccurrences,
  insertSmartField,
  linkSource,
  loadDvhDocs,
  setActiveDvhDocs,
} from '../src/renderer/dvh-smart-data'
import { createDocsDvhActions, docsSagaParticipant } from '../src/renderer/dvh-actions'

const editors: Editor[] = []
afterEach(() => {
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
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Dự án: ' }] }],
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
        ],
      },
    },
    null,
  )
  editor.commands.setTextSelection(8)
  insertSmartField(editor, dvh, dvh.model!.fields[0]!)
  const host = { editor: () => editor, filePath: () => null, readSource: async () => null }
  const registry = createDocsDvhActions(host)
  const docId = dvh.model!.docId
  const skill = createDvhScriptSkill({
    registry: () => registry,
    docId: () => docId,
    participants: () => new Map([[docId, docsSagaParticipant(host, registry, docId)]]),
    confirm: () => true,
    save: async (workflow) => {
      await registry.run(
        'Workflow.Save',
        { workflow },
        { caller: 'ai', docId, permissions: new Set(['read', 'write']) },
      )
    },
  })
  const call = (input: Record<string, unknown>) =>
    skill.executeTool({ id: 'c', name: 'propose_script', input })
  return {
    dvh,
    editor,
    registry,
    docId,
    call,
    text: () => fieldOccurrences(editor.state.doc)[0]!.text,
  }
}

describe('DVH-Script in Docs', () => {
  it('the agent checks, saves and runs scripts against the Docs catalog', async () => {
    const { dvh, registry, docId, call, text } = await setup()
    const wrong = await call({
      summary: 's',
      mode: 'run',
      source: 'Data.SetField(field:="Project.Name", valu:="x")',
    })
    expect(wrong.isError).toBe(true)
    expect(wrong.output).toContain('line 1: Data.SetField has no argument valu')

    const source = [
      'Workflow "Đặt tên"',
      'Id "wf_name"',
      'Param name As Text = "Cầu Rạch Miễu"',
      '',
      'Data.SetField(field:="Project.Name", value:=UPPER(name))',
    ].join('\n')
    const saved = await call({ summary: 'lưu', mode: 'save', source })
    expect(JSON.parse(saved.output)).toMatchObject({ ok: true, saved: 'wf_name' })
    const stored = loadWorkflows(dvh.model!).workflows[0]!
    expect(printScript(stored)).toBe(`${source}\n`)

    const ran = await call({
      summary: 'chạy',
      mode: 'run',
      source,
      params: { name: 'Cầu Mỹ Thuận' },
    })
    expect(JSON.parse(ran.output)).toMatchObject({ ok: true, changes: 1 })
    expect(text()).toBe('CẦU MỸ THUẬN')
    const sources = dvh.pending.filter((c) => c.action === 'Data.SetField').map((c) => c.source)
    expect(sources).toEqual(['script'])

    // the saved workflow runs by name too
    await registry.run(
      'Workflow.Run',
      { workflow: 'Đặt tên' },
      { caller: 'ui', docId, permissions: new Set(['read', 'write']) },
    )
    expect(text()).toBe('CẦU RẠCH MIỄU')
  })

  it('a workflow that runs itself stops at the nesting limit and changes nothing', async () => {
    const { registry, docId, call, text } = await setup()
    const loop = parseScript(
      [
        'Workflow "Vòng"',
        'Id "wf_loop"',
        '',
        'Data.SetField(field:="Project.Name", value:="X")',
        'Workflow.Run(workflow:="wf_loop")',
      ].join('\n'),
    )
    expect(loop.ok).toBe(true)
    if (!loop.ok) return
    await call({ summary: 'lưu', mode: 'save', workflow: loop.workflow })
    await expect(
      registry.run(
        'Workflow.Run',
        { workflow: 'wf_loop' },
        { caller: 'ui', docId, permissions: new Set(['read', 'write']) },
      ),
    ).rejects.toThrow(/nest deeper than 8/)
    expect(text()).toBe('A0')
  })
})

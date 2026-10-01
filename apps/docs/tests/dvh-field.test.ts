/**
 * Spike S2: a DVH Smart Field (run-level content control `dvh:f:<id>`) travels
 * run → dvhField mark → run, so a regenerated paragraph keeps the control.
 */
import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import type { Run } from '@genoffice/docx-engine'
import { inlineToRuns, runsToInline } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const SDT_PR =
  '<w:sdtPr><w:alias w:val="Project.Name"/><w:tag w:val="dvh:f:f_project"/><w:id w:val="1001"/>' +
  '<w:dataBinding w:prefixMappings="xmlns:dvh=\'urn:dvh-office:model:1\'" w:xpath="/dvh:model[1]/dvh:fields[1]/dvh:f[1]" ' +
  'w:storeItemID="{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A01}"/><w:text/></w:sdtPr>'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function editorFrom(runs: Run[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: runsToInline(runs) }],
    } as never,
  })
  editors.push(editor)
  return editor
}

const runsOf = (editor: Editor): Run[] => {
  const paragraph = editor.state.doc.firstChild!
  const nodes: Parameters<typeof inlineToRuns>[0] = []
  paragraph.forEach((child) => nodes.push(child.toJSON()))
  return inlineToRuns(nodes)
}

const fieldText = (runs: Run[]) =>
  runs
    .filter((r) => r.sdtFieldXml === SDT_PR)
    .map((r) => r.text)
    .join('')
const plainText = (runs: Run[]) =>
  runs
    .filter((r) => !r.sdtFieldXml)
    .map((r) => r.text)
    .join('')

describe('dvhField mark', () => {
  const start: Run[] = [
    { text: 'Tên dự án: ' },
    { text: 'Dự án A', sdtFieldXml: SDT_PR },
    { text: '.' },
  ]

  it('round-trips the control through the editor unchanged', () => {
    const runs = runsOf(editorFrom(start))
    expect(fieldText(runs)).toBe('Dự án A')
    expect(plainText(runs)).toBe('Tên dự án: .')
  })

  it('typing before the field changes only the surrounding text', () => {
    const editor = editorFrom(start)
    editor.commands.insertContentAt(1, 'Ghi chú – ')
    const runs = runsOf(editor)
    expect(fieldText(runs)).toBe('Dự án A')
    expect(plainText(runs)).toBe('Ghi chú – Tên dự án: .')
  })

  it('typing right after the field stays outside it (non-inclusive)', () => {
    const editor = editorFrom(start)
    // paragraph open (1) + "Tên dự án: " (11) + "Dự án A" (7)
    const end = 1 + 11 + 7
    editor.chain().setTextSelection(end).insertContent(' (mới)').run()
    const runs = runsOf(editor)
    expect(fieldText(runs)).toBe('Dự án A')
    expect(plainText(runs)).toContain(' (mới)')
  })

  it('typing inside the field keeps the new text in the control', () => {
    const editor = editorFrom(start)
    const insideField = 1 + 11 + 3 // after "Dự "
    editor.chain().setTextSelection(insideField).insertContent('X').run()
    expect(fieldText(runsOf(editor))).toBe('Dự Xán A')
  })
})

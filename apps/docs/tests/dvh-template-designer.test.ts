/**
 * P6 Template Designer: sections and column controls made in the editor are
 * written by the real save path, and the saved template generates.
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { buildBlankDocx, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { emptyModel, type DvhModel } from '@genoffice/dvh-model'
import { fillTemplate } from '@genoffice/dvh-template'
import { pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { loadDvhDocs, setActiveDvhDocs, type DvhDocsState } from '../src/renderer/dvh-smart-data'
import {
  insertColumnControl,
  saveCondition,
  sectionTagAt,
  wrapInCondition,
  wrapInRepeat,
} from '../src/renderer/dvh-template-designer'

const editors: Editor[] = []
afterEach(() => {
  setActiveDvhDocs(null)
  for (const e of editors.splice(0)) e.destroy()
})

const data = (): DvhModel => ({
  ...emptyModel('doc_template0000001'),
  collections: [
    {
      id: 'c_items',
      name: 'Items',
      columns: [
        { id: 'k_code', key: 'code', title: 'Mã', type: 'text' },
        { id: 'k_qty', key: 'qty', title: 'KL', type: 'number' },
      ],
      rows: [
        ['AB.1', 5],
        ['AB.2', 0],
      ],
    },
  ],
})

async function setup() {
  const parsed = await parseDocx(await buildBlankDocx())
  const dvh: DvhDocsState = await loadDvhDocs(parsed as never)
  setActiveDvhDocs(dvh)
  dvh.model = data()
  dvh.modelStoreItemId = '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}'
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'Biên bản' }] },
        { type: 'docParagraph', content: [{ type: 'text', text: 'Mục: ' }] },
        { type: 'docParagraph', content: [{ type: 'text', text: 'Còn thiếu khối lượng!' }] },
      ],
    } as never,
  })
  editors.push(editor)
  return { editor, dvh, parsed }
}

const paragraphPos = (editor: Editor, index: number) => {
  let pos = 0
  for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
  return pos + 1
}

describe('Template Designer', () => {
  it('makes sections and column controls that the save keeps and the engine fills', async () => {
    const { editor, dvh, parsed } = await setup()
    // a repeat over Items: paragraph 1 with the code column
    editor.commands.setTextSelection(paragraphPos(editor, 1) + 5)
    expect(insertColumnControl(editor, dvh, 'c_items', 'k_code')).toBe(true)
    editor.commands.setTextSelection(paragraphPos(editor, 1))
    expect(wrapInRepeat(editor, dvh, 'c_items')).toBe(true)
    expect(sectionTagAt(editor, 1)).toBe('dvh:repeat:c_items')
    // a condition on the summary paragraph
    const condition = saveCondition(dvh, 'COUNT >= 1', 'Có biên bản')
    expect('id' in condition).toBe(true)
    expect(saveCondition(dvh, 'COUNT >', 'sai')).toEqual({ error: 'unexpected end' })
    editor.commands.setTextSelection(paragraphPos(editor, 2))
    expect(wrapInCondition(editor, condition as never)).toBe(true)
    // nested sections are refused
    expect(wrapInCondition(editor, condition as never)).toBe(false)

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain('w:val="dvh:repeat:c_items"')
    expect(xml).toContain('w:val="dvh:c:k_code"')
    expect(xml).toContain(`w:val="dvh:if:${(condition as { id: string }).id}"`)

    const filled = await fillTemplate(saved, { model: dvh.model! })
    expect(filled.warnings).toEqual([])
    const out = await (
      await JSZip.loadAsync(filled.bytes)
    )
      .file('word/document.xml')!
      .async('string')
    expect(out).toContain('AB.1')
    expect(out).toContain('AB.2')
    expect(out.match(/Mục: /g)).toHaveLength(2)
    expect(out).toContain('Còn thiếu khối lượng!')
    expect(out).not.toMatch(/dvh:(repeat|if|c):/)
  })
})

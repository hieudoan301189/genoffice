import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  dvhFieldId,
  dvhTableId,
  parseDocx,
  saveDocx,
  wrapDvhTable,
  type Run,
  type SaveBlock,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

// Spike S2: a DVH Smart Field (run-level w:sdt tagged dvh:f:<id>, bound to the
// DVH model part) must survive DVH Office regenerating its paragraph.

const PR = (id: string, xpathIndex: number) =>
  '<w:sdtPr>' +
  `<w:alias w:val="Field ${id}"/><w:tag w:val="dvh:f:${id}"/><w:id w:val="${1000 + xpathIndex}"/>` +
  `<w:dataBinding w:prefixMappings="xmlns:dvh='urn:dvh-office:model:1'" ` +
  `w:xpath="/dvh:model[1]/dvh:fields[1]/dvh:f[${xpathIndex}]" ` +
  'w:storeItemID="{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A01}"/>' +
  '<w:text/></w:sdtPr>'
const field = (id: string, index: number, runs: string) =>
  `<w:sdt>${PR(id, index)}<w:sdtContent>${runs}</w:sdtContent></w:sdt>`
const r = (text: string, bold = false) =>
  `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`

const FIELD_P = `<w:p>${r('Tên dự án: ')}${field('f_project', 1, r('Dự án A'))}${r(' – hết.')}</w:p>`
const MULTI_RUN_P = `<w:p>${field('f_multi', 1, r('Gói ', true) + r('số 7'))}</w:p>`
const ADJACENT_P = `<w:p>${field('f_a', 1, r('A'))}${field('f_b', 2, r('B'))}</w:p>`
const OTHER_SDT_P =
  '<w:p><w:sdt><w:sdtPr><w:tag w:val="other"/><w:text/></w:sdtPr><w:sdtContent>' +
  r('khác') +
  '</w:sdtContent></w:sdt></w:p>'

async function documentXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

/** Saves `bytes` with block 0 regenerated from `edit(runs)`; every other block stays original. */
async function regenerateFirst(
  bytes: Uint8Array,
  edit: (runs: Run[]) => Run[],
): Promise<Uint8Array> {
  const doc = await parseDocx(bytes)
  const visible = doc.blocks.filter((b) => !b.hidden)
  const blocks: SaveBlock[] = visible.map((b, i) =>
    i === 0
      ? { kind: 'generated', block: { type: 'paragraph', runs: edit(b.runs ?? []) } }
      : { kind: 'original', docxIndex: b.docxIndex! },
  )
  return saveDocx(doc, blocks)
}

describe('DVH Smart Field runs', () => {
  it('parse marks the runs inside a dvh:f control with its sdtPr', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: FIELD_P }))
    const runs = doc.blocks.find((b) => !b.hidden)!.runs!
    const inField = runs.filter((run) => run.sdtFieldXml)
    expect(inField.map((run) => run.text)).toEqual(['Dự án A'])
    expect(dvhFieldId(inField[0]!.sdtFieldXml!)).toBe('f_project')
    expect(inField[0]!.sdtFieldXml).toContain('w:dataBinding')
    expect(
      runs
        .filter((run) => !run.sdtFieldXml)
        .map((run) => run.text)
        .join(''),
    ).toBe('Tên dự án:  – hết.')
  })

  it('an untouched paragraph still saves byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: FIELD_P })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(
      doc,
      doc.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(Buffer.from(saved).equals(Buffer.from(bytes))).toBe(true)
  })

  it('editing text around the field regenerates the same control around the field text', async () => {
    const saved = await regenerateFirst(await buildDocx({ bodyXml: FIELD_P }), (runs) =>
      runs.map((run, i) => (i === 0 ? { ...run, text: 'Tên công trình: ' } : run)),
    )
    const xml = await documentXml(saved)
    expect(xml).toContain('Tên công trình: ')
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1)
    expect(xml).toContain(`<w:sdt>${PR('f_project', 1)}<w:sdtContent>`)
    expect(xml).toMatch(
      /<w:sdtContent><w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>Dự án A<\/w:t><\/w:r><\/w:sdtContent><\/w:sdt>/,
    )
    // and the regenerated file reads back as a field again
    const again = (await parseDocx(saved)).blocks.find((b) => !b.hidden)!.runs!
    expect(again.find((run) => run.sdtFieldXml)?.text).toBe('Dự án A')
  })

  it('editing the field text keeps the control around the new text', async () => {
    const saved = await regenerateFirst(await buildDocx({ bodyXml: FIELD_P }), (runs) =>
      runs.map((run) => (run.sdtFieldXml ? { ...run, text: 'Dự án B' } : run)),
    )
    const xml = await documentXml(saved)
    expect(xml).toMatch(
      /<w:sdtContent><w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>Dự án B<\/w:t><\/w:r><\/w:sdtContent>/,
    )
  })

  it('a field holding several formatted runs regenerates as one control', async () => {
    const saved = await regenerateFirst(await buildDocx({ bodyXml: MULTI_RUN_P }), (runs) => runs)
    const xml = await documentXml(saved)
    expect(xml.match(/<w:sdt>/g)).toHaveLength(1)
    expect(xml).toMatch(
      /<w:sdtContent><w:r><w:rPr><w:b\/>[\s\S]*?Gói [\s\S]*?số 7[\s\S]*?<\/w:sdtContent>/,
    )
  })

  it('adjacent fields stay separate controls', async () => {
    const saved = await regenerateFirst(await buildDocx({ bodyXml: ADJACENT_P }), (runs) => runs)
    const xml = await documentXml(saved)
    expect(xml.match(/<w:sdt>/g)).toHaveLength(2)
    expect(xml).toContain('w:val="dvh:f:f_a"')
    expect(xml).toContain('w:val="dvh:f:f_b"')
  })

  it('other inline controls keep the previous behavior (flattened on regenerate)', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: OTHER_SDT_P }))
    const runs = doc.blocks.find((b) => !b.hidden)!.runs!
    expect(runs.some((run) => run.sdtFieldXml)).toBe(false)
  })
})

// P2: a rendered DVH.Table is a docx table inside a block control tagged dvh:t:<id>
describe('DVH table controls', () => {
  const TABLE_PR =
    '<w:sdtPr><w:alias w:val="Bảng KL"/><w:tag w:val="dvh:t:t_workitems"/><w:id w:val="77"/></w:sdtPr>'
  const tbl =
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' +
    `<w:tr><w:tc><w:p>${r('Mã')}</w:p></w:tc><w:tc><w:p>${r('Khối lượng')}</w:p></w:tc></w:tr>` +
    `<w:tr><w:tc><w:p>${r('AB.1')}</w:p></w:tc><w:tc><w:p>${r('120,50')}</w:p></w:tc></w:tr></w:tbl>`
  const wrapped = `<w:sdt>${TABLE_PR}<w:sdtContent>${tbl}</w:sdtContent></w:sdt>`

  it('parse gives the table block its dvh:t sdtPr; other table controls get none', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: wrapped + `<w:p>${r('x')}</w:p>` }))
    const table = doc.blocks.find((b) => b.type === 'table')!
    expect(table.tableSdtPr).toBe(TABLE_PR)
    expect(dvhTableId(table.tableSdtPr!)).toBe('t_workitems')
    const other = await parseDocx(
      await buildDocx({ bodyXml: wrapped.replace('dvh:t:t_workitems', 'report-table') }),
    )
    expect(other.blocks.find((b) => b.type === 'table')!.tableSdtPr).toBeUndefined()
  })

  it('an untouched table keeps its control byte for byte; wrapDvhTable rewraps generated XML once', async () => {
    const bytes = await buildDocx({ bodyXml: wrapped })
    const doc = await parseDocx(bytes)
    const blocks: SaveBlock[] = doc.blocks
      .filter((b) => !b.hidden)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    expect(await documentXml(await saveDocx(doc, blocks))).toContain(wrapped)
    expect(wrapDvhTable(tbl, TABLE_PR)).toBe(wrapped)
    expect(wrapDvhTable(wrapped, TABLE_PR)).toBe(wrapped)
  })
})

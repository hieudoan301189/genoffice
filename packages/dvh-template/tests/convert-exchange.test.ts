/**
 * P6: DVH-Tool templates become Smart Templates (fields, repeating rows and
 * blocks, model part created), the converted template generates documents,
 * and QLCL data round-trips through its workbook layout and the JSON format.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { MODEL_NS, parseModelXml } from '@genoffice/dvh-model'
import {
  convertDvhToolTemplate,
  exchangeJsonSchema,
  exportExchangeJson,
  exportQlclWorkbook,
  fillTemplate,
  findItemByNamespace,
  importExchangeJson,
  importQlclWorkbook,
  keyOf,
  readWorkbookGrids,
  replacePlaceholders,
} from '../src/index'
import { cell, p, row, run, X } from './fixtures'

const W = 'xmlns:w="w" xmlns:r="r"'

async function dvhToolDocx(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  )
  const body = [
    // Word split the placeholder over three runs after an edit
    p(run('Công trình: '), run('<<Ten'), run('Cong', true), run('Trinh>> — chủ đầu tư <<CDT>>.')),
    p(run('Lỗi chèn <<Ngắt'), '<w:r><w:tab/></w:r>', run('dòng>>')),
    `<w:tbl><w:tblPr/>` +
      row(cell(p(run('STT'))), cell(p(run('Công việc'))), cell(p(run('Khối lượng')))) +
      row(cell(p(run('BD_Bang_KhoiLuong')))) +
      row(cell(p(run('<<STT>>'))), cell(p(run('<<Công việc>>'))), cell(p(run('<<Khối lượng>>')))) +
      row(cell(p(run('KT_Bang_KhoiLuong')))) +
      row(cell(p(run('Tổng'))), cell(p()), cell(p(run('<<TongKL>>')))) +
      '</w:tbl>',
    p(run('BD_Bang_ThanhPhan')),
    p(run('- Ông/bà <<Họ tên>>, chức vụ <<Chức vụ>>')),
    p(run('KT_Bang_ThanhPhan')),
    '<w:sectPr/>',
  ].join('')
  zip.file('word/document.xml', `${X}<w:document ${W}><w:body>${body}</w:body></w:document>`)
  return zip.generateAsync({ type: 'uint8array' })
}

describe('DVH-Tool conversion', () => {
  it('finds placeholders split over runs, keeping the formatting of the first run', () => {
    const out = replacePlaceholders(
      p(run('A <<X'), run('Y>> B', true)),
      (name, rPr) => `[${name}|${rPr ? 'fmt' : 'plain'}]`,
    )
    expect(out.xml).toBe(
      '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r>[XY|plain]<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> B</w:t></w:r></w:p>',
    )
    expect(keyOf('Khối lượng (m3)')).toBe('khoi_luong_m3')
    expect(keyOf('Đơn giá')).toBe('don_gia')
  })

  it('turns <<Field>> into Smart Fields and BD_Bang…KT_Bang into repeating rows/blocks', async () => {
    const { bytes, model, report } = await convertDvhToolTemplate(await dvhToolDocx())
    expect(report.fields).toEqual(['TenCongTrinh', 'CDT', 'TongKL'])
    expect(report.regions).toEqual([
      { name: 'KhoiLuong', columns: ['STT', 'Công việc', 'Khối lượng'], kind: 'rows' },
      { name: 'ThanhPhan', columns: ['Họ tên', 'Chức vụ'], kind: 'blocks' },
    ])
    expect(report.warnings).toEqual(['<<Ngắt dòng>> spans a tab, break or picture: left as text'])
    const zip = await JSZip.loadAsync(bytes)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).not.toContain('BD_Bang')
    expect(xml).not.toContain('<<TenCongTrinh>>')
    expect(xml.match(/w:val="dvh:f:/g)).toHaveLength(3)
    expect(xml).toContain(`w:val="dvh:repeatrows:${model.collections[0]!.id}"`)
    expect(xml).toContain(`w:val="dvh:repeat:${model.collections[1]!.id}"`)
    // the model part was created and is found by namespace, with its plumbing
    const modelPath = await findItemByNamespace(zip, MODEL_NS)
    expect(parseModelXml(await zip.file(modelPath!)!.async('string')).fields).toHaveLength(3)
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(
      'customXml/item1.xml',
    )
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('itemProps1.xml')

    // the converted template generates
    model.fields[0]!.value = 'Cầu Bến Thủy 3'
    model.collections[0]!.rows = [
      [1, 'Đào móng', 12.5],
      [2, 'Bê tông lót', 3],
    ]
    model.collections[1]!.rows = [['Nguyễn Văn A', 'Giám sát']]
    const filled = await fillTemplate(bytes, { model })
    const out = await (
      await JSZip.loadAsync(filled.bytes)
    )
      .file('word/document.xml')!
      .async('string')
    expect(out).toContain('Cầu Bến Thủy 3')
    expect(out).toContain('Đào móng')
    expect(out).toContain('Bê tông lót')
    expect(out).toContain('Nguyễn Văn A')
    expect(out).toContain('Tổng')
  })
})

describe('QLCL exchange', () => {
  it('round-trips Smart Data through the QLCL workbook layout, keeping ids on re-import', async () => {
    const qlcl = await exportQlclWorkbook({
      docId: 'doc_aaaaaaaaaaaaaaaa',
      schemaVersion: 1,
      fields: [
        { id: 'f_a', name: 'TenCongTrinh', type: 'text', value: 'Cầu & cống', access: 'readwrite' },
        { id: 'f_b', name: 'SoHopDong', type: 'number', value: 12, access: 'readwrite' },
      ],
      collections: [
        {
          id: 'c_x',
          name: 'Data',
          columns: [
            { id: 'k1', key: 'ma', title: 'Mã', type: 'text' },
            { id: 'k2', key: 'kl', title: 'KL', type: 'number' },
            { id: 'k3', key: 'dat', title: 'Đạt', type: 'boolean' },
          ],
          rows: [
            ['AB.1', 1.5, true],
            ['AB.2', null, false],
          ],
        },
      ],
      tables: [],
      links: [],
    })
    const grids = await readWorkbookGrids(qlcl)
    expect([...grids.keys()]).toEqual(['ThongTin', 'Data'])
    const imported = await importQlclWorkbook(qlcl)
    expect(imported.fields.map((f) => [f.name, f.value])).toEqual([
      ['TenCongTrinh', 'Cầu & cống'],
      ['SoHopDong', 12],
    ])
    expect(imported.collections[0]!.rows).toEqual([
      ['AB.1', 1.5, true],
      ['AB.2', null, false],
    ])
    expect(imported.collections[0]!.columns.map((c) => [c.key, c.type])).toEqual([
      ['ma', 'text'],
      ['kl', 'number'],
      ['dat', 'boolean'],
    ])
    // a second import into the same model keeps every id
    const again = await importQlclWorkbook(qlcl, imported)
    expect(again.fields.map((f) => f.id)).toEqual(imported.fields.map((f) => f.id))
    expect(again.collections[0]!.id).toBe(imported.collections[0]!.id)
    expect(again.collections[0]!.columns.map((c) => c.id)).toEqual(
      imported.collections[0]!.columns.map((c) => c.id),
    )
    // deterministic export
    expect(await exportQlclWorkbook(again)).toEqual(await exportQlclWorkbook(again))
  })

  it('the JSON exchange format is versioned and validated', () => {
    const json = exportExchangeJson(
      { docId: 'doc_x', schemaVersion: 1, fields: [], collections: [], tables: [], links: [] },
      '2026-10-02T00:00:00.000Z',
    )
    expect(importExchangeJson(json).docId).toBe('doc_x')
    expect(() => importExchangeJson('{"format":"other"}')).toThrow(/not a dvh-exchange v1/)
    const schema = exchangeJsonSchema() as { properties: Record<string, unknown> }
    expect(Object.keys(schema.properties)).toEqual([
      'format',
      'version',
      'producer',
      'exportedAt',
      'model',
    ])
  })
})

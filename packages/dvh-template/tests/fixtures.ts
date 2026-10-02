/** A small Smart Template docx, built the way Word stores one. */
import JSZip from 'jszip'
import {
  fieldSdtPrXml,
  itemPropsXml,
  MODEL_NS,
  serializeModelXml,
  type DvhModel,
} from '@genoffice/dvh-model'

export const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
export const STORE = '{11111111-2222-3333-4444-555555555555}'
const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

export const run = (text: string, bold = false) =>
  `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</w:t></w:r>`
export const sdt = (tag: string, content: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`
export const p = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`
export const field = (id: string, name: string, shown: string) =>
  `<w:sdt>${fieldSdtPrXml({ fieldId: id, alias: name, sdtId: 7, storeItemId: STORE })}<w:sdtContent>${run(shown, true)}</w:sdtContent></w:sdt>`
export const cell = (...paragraphs: string[]) =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${paragraphs.join('')}</w:tc>`
export const row = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`

export const workModel = (count: number): DvhModel => ({
  docId: 'doc_template000000001',
  schemaVersion: 1,
  fields: [
    {
      id: 'f_project',
      name: 'Project.Name',
      type: 'text',
      value: 'Cầu Bến Thủy 3',
      access: 'readwrite',
    },
    { id: 'f_logo', name: 'Project.Logo', type: 'image', value: 'logo.png', access: 'readwrite' },
  ],
  collections: [
    {
      id: 'c_work',
      name: 'WorkItems',
      columns: [
        { id: 'k_code', key: 'code', title: 'Mã', type: 'text' },
        { id: 'k_name', key: 'name', title: 'Công việc', type: 'text' },
        { id: 'k_qty', key: 'qty', title: 'Khối lượng', type: 'number', numFmt: '#,##0.00' },
        { id: 'k_ok', key: 'ok', title: 'Đạt', type: 'boolean' },
      ],
      rows: Array.from({ length: count }, (_, i) => [
        `AB.${i + 1}`,
        `Hạng mục ${i + 1} & <phụ>`,
        1000 + i * 12.5,
        i % 3 !== 0,
      ]),
    },
    {
      id: 'c_checks',
      name: 'Checks',
      columns: [
        { id: 'k_item', key: 'item', title: 'Nội dung', type: 'text' },
        { id: 'k_res', key: 'result', title: 'Kết quả', type: 'text' },
      ],
      rows: [
        ['Cao độ', 'Đạt'],
        ['Kích thước', 'Đạt'],
        ['Vệ sinh', 'Không đạt'],
      ],
    },
  ],
  tables: [
    {
      id: 't_checks',
      name: 'Checks table',
      source: { collectionId: 'c_checks' },
      columns: [
        { id: 'tc_1', columnId: 'k_item', title: 'Nội dung' },
        { id: 'tc_2', columnId: 'k_res', title: 'Kết quả' },
      ],
      style: { mode: 'destination' },
      layout: { repeatHeader: true, keepRowsTogether: false, widths: 'page' },
    },
  ],
  links: [],
  conditions: [
    { id: 'cond_fail', expr: 'NOT row.ok', name: 'Không đạt' },
    { id: 'cond_big', expr: 'row.qty >= 1100', name: 'Khối lượng lớn' },
    { id: 'cond_failed_checks', expr: 'row.result <> "Đạt"' },
  ],
})

/** One acceptance record (biên bản nghiệm thu) per work item. */
export function templateDocument(): string {
  const body = [
    p(run('BIÊN BẢN NGHIỆM THU — '), field('f_project', 'Project.Name', 'Tên dự án')),
    p(
      run('Công việc: '),
      sdt('dvh:c:k_code', run('mã')),
      run(' — '),
      sdt('dvh:c:k_name', run('tên')),
    ),
    p(run('Khối lượng: '), sdt('dvh:c:k_qty', run('0'))),
    sdt('dvh:if:cond_fail', p(run('Kết luận: KHÔNG ĐẠT, yêu cầu sửa chữa.'))),
    sdt('dvh:if:cond_big', p(run('Ghi chú: khối lượng lớn.'))),
    sdt(
      'dvh:repeatrows:c_checks',
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
        row(cell(p(run('Nội dung', true))), cell(p(run('Kết quả', true)))) +
        row(cell(p(sdt('dvh:c:k_item', run('x')))), cell(p(sdt('dvh:c:k_res', run('y'))))) +
        row(cell(p(run('Hết bảng'))), cell(p())) +
        '</w:tbl>',
    ),
    sdt('dvh:repeat:c_checks:cond_failed_checks', p(run('Lỗi: '), sdt('dvh:c:k_item', run('?')))),
    sdt('dvh:t:t_checks', '<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>'),
    p(
      sdt(
        'dvh:img:f_logo',
        '<w:r><w:drawing><wp:inline xmlns:wp="wp"><a:graphic xmlns:a="a"><a:graphicData><pic:pic xmlns:pic="pic"><pic:blipFill><a:blip r:embed="rIdLogo"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>',
      ),
    ),
    '<w:sectPr/>',
  ].join('')
  return `${X}<w:document ${W}><w:body>${body}</w:body></w:document>`
}

export async function templateDocx(model: DvhModel = workModel(3)): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/></Types>`,
  )
  zip.file('word/document.xml', templateDocument())
  zip.file(
    'word/_rels/document.xml.rels',
    `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLogo" Type="image" Target="media/image1.png"/><Relationship Id="rId9" Type="customXml" Target="../customXml/item1.xml"/></Relationships>`,
  )
  zip.file('word/media/image1.png', new Uint8Array([1, 2, 3]))
  zip.file('customXml/item1.xml', serializeModelXml(model))
  zip.file('customXml/itemProps1.xml', itemPropsXml(STORE, MODEL_NS))
  return zip.generateAsync({ type: 'uint8array' })
}

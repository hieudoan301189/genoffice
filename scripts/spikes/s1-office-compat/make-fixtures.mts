// Spike S1: builds the probe documents. The docx starts from DVH Office's own
// blank template (docx-engine); the xlsx is hand-assembled so every part under
// test is explicit. Usage: tsx make-fixtures.mts <outDir>
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { buildBlankDocx } from '@genoffice/docx-engine'
import {
  CUSTOM_XML_PROPS_TYPE,
  CUSTOM_XML_REL,
  DATE_VALUE,
  DEFINED_FIELD,
  DEFINED_TABLE,
  DOCX_DOC_ID,
  FIELD_DATE,
  FIELD_PROJECT,
  HISTORY_NS,
  HISTORY_STORE_ID,
  MODEL_NS,
  MODEL_STORE_ID,
  PREFIX_MAPPINGS,
  PROJECT_VALUE,
  STALE_A,
  STALE_B,
  TABLE_ITEMS,
  WORD_EDIT_VALUE,
  WORK_ITEMS,
  XLSX_DOC_ID,
  XPATH_BY_ID,
  XPATH_BY_POSITION,
  escapeXml,
  historyPartXml,
  itemPropsXml,
  itemRelsXml,
  modelPartXml,
} from './dvh-ooxml.mts'

const outDir = process.argv[2] ?? join(import.meta.dirname, 'out')
mkdirSync(outDir, { recursive: true })

const SMALL_HISTORY = 20_000
const BIG_HISTORY = 5_000_000

function addCustomXmlParts(zip: JSZip, docId: string, historyBytes: number): void {
  zip.file('customXml/item1.xml', modelPartXml(docId))
  zip.file('customXml/itemProps1.xml', itemPropsXml(MODEL_STORE_ID, MODEL_NS))
  zip.file('customXml/_rels/item1.xml.rels', itemRelsXml('itemProps1.xml'))
  zip.file('customXml/item2.xml', historyPartXml(docId, historyBytes))
  zip.file('customXml/itemProps2.xml', itemPropsXml(HISTORY_STORE_ID, HISTORY_NS))
  zip.file('customXml/_rels/item2.xml.rels', itemRelsXml('itemProps2.xml'))
}

function contentTypeOverrides(): string {
  return [1, 2]
    .map(
      (n) =>
        `<Override PartName="/customXml/itemProps${n}.xml" ContentType="${CUSTOM_XML_PROPS_TYPE}"/>`,
    )
    .join('')
}

const run = (text: string, rPr = ''): string =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`

function boundField(
  id: string,
  alias: string,
  sdtId: number,
  xpath: string,
  shown: string,
): string {
  return (
    '<w:sdt><w:sdtPr>' +
    `<w:alias w:val="${alias}"/><w:tag w:val="dvh:f:${id}"/><w:id w:val="${sdtId}"/>` +
    `<w:dataBinding w:prefixMappings="${PREFIX_MAPPINGS}" w:xpath="${xpath}" w:storeItemID="${MODEL_STORE_ID}"/>` +
    '<w:text/></w:sdtPr>' +
    `<w:sdtContent>${run(shown)}</w:sdtContent></w:sdt>`
  )
}

function tableCell(text: string, width: number, header: boolean): string {
  const shading = header ? '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' : ''
  return (
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${shading}</w:tcPr>` +
    `<w:p>${run(text, header ? '<w:b/>' : '')}</w:p></w:tc>`
  )
}

function workItemsTable(): string {
  const widths = [1500, 4500, 2000]
  const border = (side: string) =>
    `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`
  const borders = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')
  const header =
    '<w:tr><w:trPr><w:tblHeader/></w:trPr>' +
    ['Mã', 'Công việc', 'Khối lượng'].map((t, i) => tableCell(t, widths[i]!, true)).join('') +
    '</w:tr>'
  const rows = WORK_ITEMS.map(
    ([code, name, qty]) =>
      '<w:tr>' +
      [code, name, qty.toFixed(3)].map((t, i) => tableCell(t, widths[i]!, false)).join('') +
      '</w:tr>',
  ).join('')
  return (
    '<w:sdt><w:sdtPr>' +
    `<w:alias w:val="WorkItems"/><w:tag w:val="dvh:t:${TABLE_ITEMS}"/><w:id w:val="1003"/>` +
    '</w:sdtPr><w:sdtContent>' +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${borders}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    header +
    rows +
    '</w:tbl></w:sdtContent></w:sdt>'
  )
}

async function buildDocx(historyBytes: number): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await buildBlankDocx())
  const body =
    `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run('Spike S1 – Smart Data trong DOCX')}</w:p>` +
    `<w:p>${run('Tên dự án: ')}${boundField(FIELD_PROJECT, 'Project.Name', 1001, XPATH_BY_ID, STALE_A)}</w:p>` +
    `<w:p>${run('Ngày nghiệm thu: ')}${boundField(FIELD_DATE, 'Acceptance.Date', 1002, XPATH_BY_POSITION, STALE_B)}</w:p>` +
    `<w:p>${run('Đoạn thường – DVH Office sẽ sửa đoạn này.')}</w:p>` +
    workItemsTable() +
    `<w:p>${run('Hết tài liệu.')}</w:p>`
  const docPath = 'word/document.xml'
  const documentXml = await zip.file(docPath)!.async('string')
  zip.file(docPath, documentXml.replace('<w:body><w:p/>', `<w:body>${body}`))

  const relsPath = 'word/_rels/document.xml.rels'
  const rels = await zip.file(relsPath)!.async('string')
  zip.file(
    relsPath,
    rels.replace(
      '</Relationships>',
      `<Relationship Id="rIdDvh1" Type="${CUSTOM_XML_REL}" Target="../customXml/item1.xml"/>` +
        `<Relationship Id="rIdDvh2" Type="${CUSTOM_XML_REL}" Target="../customXml/item2.xml"/>` +
        '</Relationships>',
    ),
  )
  const typesPath = '[Content_Types].xml'
  const types = await zip.file(typesPath)!.async('string')
  zip.file(typesPath, types.replace('</Types>', `${contentTypeOverrides()}</Types>`))
  addCustomXmlParts(zip, DOCX_DOC_ID, historyBytes)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const XLSX_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.000"/></numFmts>' +
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>' +
  '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="3"><fill><patternFill patternType="none"/></fill>' +
  '<fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '</styleSheet>'

const str = (ref: string, text: string, style = 0): string =>
  `<c r="${ref}"${style ? ` s="${style}"` : ''} t="inlineStr"><is><t>${escapeXml(text)}</t></is></c>`

function sheetXml(): string {
  const rows: string[] = [
    `<row r="1">${str('A1', 'Spike S1 – Smart Data trong XLSX')}</row>`,
    `<row r="5">${str('A5', 'Tên dự án')}${str('B5', PROJECT_VALUE)}` +
      `<c r="C5" t="str"><f>${DEFINED_FIELD}</f><v>${escapeXml(PROJECT_VALUE)}</v></c>` +
      `<c r="D5" t="str"><f>TenDuAn</f><v>${escapeXml(PROJECT_VALUE)}</v></c></row>`,
    `<row r="6">${str('A6', 'Ngày nghiệm thu')}${str('B6', DATE_VALUE)}</row>`,
    `<row r="8">${['Mã', 'Công việc', 'Khối lượng'].map((t, i) => str(`${'ABC'[i]}8`, t, 1)).join('')}</row>`,
    ...WORK_ITEMS.map(
      ([code, name, qty], i) =>
        `<row r="${9 + i}">${str(`A${9 + i}`, code)}${str(`B${9 + i}`, name)}` +
        `<c r="C${9 + i}" s="2"><v>${qty}</v></c></row>`,
    ),
  ]
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<cols><col min="1" max="1" width="16" customWidth="1"/><col min="2" max="4" width="28" customWidth="1"/></cols>' +
    `<sheetData>${rows.join('')}</sheetData></worksheet>`
  )
}

async function buildXlsx(historyBytes: number): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      contentTypeOverrides() +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
      '<definedNames>' +
      `<definedName name="${DEFINED_FIELD}" hidden="1">Sheet1!$B$5</definedName>` +
      `<definedName name="${DEFINED_TABLE}" hidden="1">Sheet1!$A$8:$C$11</definedName>` +
      '<definedName name="TenDuAn">Sheet1!$B$5</definedName>' +
      '</definedNames><calcPr calcId="191029"/></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      `<Relationship Id="rId3" Type="${CUSTOM_XML_REL}" Target="../customXml/item1.xml"/>` +
      `<Relationship Id="rId4" Type="${CUSTOM_XML_REL}" Target="../customXml/item2.xml"/>` +
      '</Relationships>',
  )
  zip.file('xl/styles.xml', XLSX_STYLES)
  zip.file('xl/worksheets/sheet1.xml', sheetXml())
  addCustomXmlParts(zip, XLSX_DOC_ID, historyBytes)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const outputs: [string, Buffer][] = [
  ['dvh-s1.docx', await buildDocx(SMALL_HISTORY)],
  ['dvh-s1-big.docx', await buildDocx(BIG_HISTORY)],
  ['dvh-s1.xlsx', await buildXlsx(SMALL_HISTORY)],
  ['dvh-s1-big.xlsx', await buildXlsx(BIG_HISTORY)],
]
for (const [name, bytes] of outputs) {
  writeFileSync(join(outDir, name), bytes)
  console.log(`${name}\t${bytes.length} bytes`)
}
// read by office-roundtrip.ps1 (kept ASCII-only, so the Vietnamese edit text travels as UTF-8 JSON)
writeFileSync(
  join(outDir, 'office-config.json'),
  JSON.stringify({
    modelNs: MODEL_NS,
    projectTag: `dvh:f:${FIELD_PROJECT}`,
    projectXPath: XPATH_BY_ID,
    wordEditValue: WORD_EDIT_VALUE,
  }),
)

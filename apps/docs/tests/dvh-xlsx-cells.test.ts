/**
 * P3: a linked workbook saved by Microsoft Excel keeps a stale DVH model part;
 * the document reads fields and collections from the cells instead.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { emptyModel, type DvhModel } from '@genoffice/dvh-model'
import { modelFromCells, parseRef } from '../src/main/dvh-xlsx-cells'

const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

/** A workbook as Excel writes it: shared strings, styles with a custom format, a theme color. */
async function excelWorkbook(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    `${X}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Khối lượng" sheetId="1" r:id="rId1"/></sheets>` +
      `<definedNames><definedName name="_dvh.c.c_work" hidden="1">'Khối lượng'!$A$1:$C$2</definedName>` +
      `<definedName name="_dvh.f.f_total" hidden="1">'Khối lượng'!$G$1</definedName></definedNames></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
  )
  zip.file(
    'xl/theme/theme1.xml',
    `${X}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x">` +
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
      '</a:clrScheme></a:themeElements></a:theme>',
  )
  zip.file(
    'xl/sharedStrings.xml',
    `${X}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Mã</t></si><si><t>Công việc</t></si>` +
      '<si><r><t>Khối </t></r><r><rPr><b/></rPr><t>lượng</t></r></si><si><t>AB.1</t></si><si><t>Đào &amp; đắp</t></si>' +
      '<si><t>AB.2</t></si><si><t>Bê tông</t></si></sst>',
  )
  zip.file(
    'xl/styles.xml',
    `${X}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00&quot; m3&quot;"/></numFmts>' +
      '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><color theme="4"/></font><font><color rgb="FFC00000"/><sz val="11"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill></fills>' +
      '<borders count="2"><border/><border><left style="thin"><color rgb="FF808080"/></left><bottom style="medium"><color rgb="FF808080"/></bottom></border></borders>' +
      '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyAlignment="1"><alignment horizontal="center"/></xf>' +
      '<xf numFmtId="164" fontId="2" fillId="0" borderId="0"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>',
  )
  // the range says A1:C2, but Excel users typed a third row and a fourth column title
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${X}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="2" max="2" width="30" customWidth="1"/></cols><sheetData>` +
      '<row r="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c><c r="C1" s="1" t="s"><v>2</v></c><c r="D1" s="1" t="inlineStr"><is><t>Ghi chú</t></is></c><c r="G1" s="3"><f>SUM(C2:C3)</f><v>138.5</v></c></row>' +
      '<row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2" t="s"><v>4</v></c><c r="C2" s="2"><v>120.5</v></c></row>' +
      '<row r="3"><c r="A3" t="s"><v>5</v></c><c r="B3" t="s"><v>6</v></c><c r="C3" s="2"><v>18</v></c><c r="D3" t="str"><f>"ok"</f><v>ok</v></c></row>' +
      '<row r="5"><c r="A5" t="inlineStr"><is><t>ngoài bảng</t></is></c></row>' +
      '</sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'uint8array' })
}

const staleModel: DvhModel = {
  ...emptyModel('doc_workbook00000001'),
  fields: [{ id: 'f_total', name: 'Total', type: 'number', value: 1, access: 'readwrite' }],
  collections: [
    {
      id: 'c_work',
      name: 'WorkItems',
      columns: [
        { id: 'col_a', key: 'code', title: 'Mã', type: 'text' },
        { id: 'col_b', key: 'name', title: 'Công việc', type: 'text' },
        { id: 'col_c', key: 'qty', title: 'Khối lượng', type: 'number' },
      ],
      rows: [['OLD', 'stale', 0]],
    },
  ],
}

describe('modelFromCells', () => {
  it('parses quoted sheet references', () => {
    expect(parseRef("'Khối lượng'!$A$1:$C$2")).toEqual({
      sheetName: 'Khối lượng',
      row: 0,
      column: 0,
      rows: 2,
      columns: 3,
    })
  })

  it('reads values, grown ranges, formats, styles and widths from the cells', async () => {
    const model = await modelFromCells(await excelWorkbook(), staleModel)
    const work = model.collections[0]!
    expect(work.columns.map((c) => c.title)).toEqual(['Mã', 'Công việc', 'Khối lượng', 'Ghi chú'])
    // existing columns keep their ids, the new one gets a fresh id
    expect(work.columns.slice(0, 3).map((c) => c.id)).toEqual(['col_a', 'col_b', 'col_c'])
    expect(work.columns[3]!.id).not.toMatch(/^col_[abc]$/)
    expect(work.rows).toEqual([
      ['AB.1', 'Đào & đắp', 120.5, null],
      ['AB.2', 'Bê tông', 18, 'ok'],
    ])
    expect(work.columns[0]!.headerStyle).toEqual({
      bold: true,
      color: '#4472C4',
      fontSize: 12,
      fill: '#FFF2CC',
      align: 'center',
      border: { color: '#808080', width: 'medium' },
    })
    expect(work.columns[2]).toMatchObject({
      type: 'number',
      numFmt: '#,##0.00" m3"',
      style: { color: '#C00000', fontSize: 11 },
    })
    expect(work.columns[1]!.width).toBe(215)
    expect(model.fields[0]!.value).toBe(138.5)
    // the model part's other content is untouched
    expect(staleModel.collections[0]!.rows).toEqual([['OLD', 'stale', 0]])
  })
})

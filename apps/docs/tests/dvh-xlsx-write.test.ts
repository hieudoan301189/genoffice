/**
 * P3 write-back into a closed workbook: the bound cell, the model part (text +
 * revision) and the history part change; a source edited since the base
 * (by DVH or by Excel) reports a conflict and nothing is written.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  HISTORY_NS,
  MODEL_NS,
  parseHistoryXml,
  parseModelXml,
  serializeModelXml,
  type DvhModel,
} from '@genoffice/dvh-model'
import { readCustomXmlPart } from '@genoffice/docx-engine'
import { modelFromCells } from '../src/main/dvh-xlsx-cells'
import { setCellInSheetXml, writeFieldsToXlsx } from '../src/main/dvh-xlsx-write'

const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const DOC = 'doc_workbook00000001'

const model: DvhModel = {
  docId: DOC,
  schemaVersion: 1,
  fields: [
    {
      id: 'f_name',
      name: 'Project.Name',
      type: 'text',
      value: 'Dự án A',
      access: 'readwrite',
      rev: 2,
    },
    { id: 'f_qty', name: 'Lot.Qty', type: 'number', value: 5, access: 'readwrite' },
    { id: 'f_sum', name: 'Lot.Total', type: 'number', value: 10, access: 'readwrite' },
    { id: 'f_new', name: 'Lot.Note', type: 'text', value: null, access: 'readwrite' },
  ],
  collections: [],
  tables: [],
  links: [],
}

/** As Excel saves it: shared strings, a formula, no history part. */
async function workbook(nameInCell = 'Dự án A'): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`,
  )
  zip.file(
    'xl/workbook.xml',
    `${X}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Info" sheetId="1" r:id="rId1"/></sheets>` +
      '<definedNames><definedName name="_dvh.f.f_name" hidden="1">Info!$B$1</definedName>' +
      '<definedName name="_dvh.f.f_qty" hidden="1">Info!$B$2</definedName>' +
      '<definedName name="_dvh.f.f_sum" hidden="1">Info!$B$3</definedName>' +
      '<definedName name="_dvh.f.f_new" hidden="1">Info!$C$5</definedName></definedNames>' +
      '<calcPr calcId="191029"/></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>',
  )
  zip.file(
    'xl/sharedStrings.xml',
    `${X}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Tên</t></si><si><t>${nameInCell}</t></si></sst>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${X}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="4" t="s"><v>1</v></c></row>' +
      '<row r="2"><c r="B2"><v>5</v></c></row>' +
      '<row r="3"><c r="B3"><f>B2*2</f><v>10</v></c></row>' +
      '<row r="7"><c r="A7"><v>1</v></c></row>' +
      '</sheetData></worksheet>',
  )
  zip.file('customXml/item1.xml', serializeModelXml(model))
  return zip.generateAsync({ type: 'uint8array' })
}

const write = (
  fieldId: string,
  value: string | number | null,
  expected: { rev: number; text: string } | null,
  force = false,
) => ({
  fieldId,
  value,
  expected,
  force,
})

describe('writeFieldsToXlsx', () => {
  it('writes the cell, the model part (text + rev) and a history change set', async () => {
    const out = await writeFieldsToXlsx(await workbook(), {
      docId: DOC,
      writes: [
        write('f_name', 'Dự án B & <C>', { rev: 2, text: 'Dự án A' }),
        write('f_qty', 7, { rev: 0, text: '5' }),
        write('f_new', 'ghi chú', null),
      ],
    })
    expect(out.results.map((r) => r.status)).toEqual(['written', 'written', 'written'])
    expect(out.results[0]!.current).toEqual({ rev: 3, text: 'Dự án B & <C>' })
    const bytes = out.bytes!
    const part = await readCustomXmlPart(bytes, MODEL_NS)
    const saved = parseModelXml(part!.xml)
    expect(saved.fields.find((f) => f.id === 'f_name')).toMatchObject({
      value: 'Dự án B & <C>',
      rev: 3,
    })
    expect(saved.fields.find((f) => f.id === 'f_qty')).toMatchObject({ value: 7, rev: 1 })
    // the cells read back the new values (C5 was created in a new row, in order)
    const cells = await modelFromCells(bytes, saved)
    expect(cells.fields.map((f) => f.value)).toEqual(['Dự án B & <C>', 7, 10, 'ghi chú'])
    const zip = await JSZip.loadAsync(bytes)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<c r="B1" s="4" t="inlineStr">')
    expect(sheet.indexOf('<row r="5">')).toBeLessThan(sheet.indexOf('<row r="7">'))
    expect(await zip.file('xl/workbook.xml')!.async('string')).toContain(
      '<calcPr calcId="191029" fullCalcOnLoad="1"/>',
    )
    // the history part was added with its plumbing
    const history = parseHistoryXml((await readCustomXmlPart(bytes, HISTORY_NS))!.xml)
    expect(history.changes).toHaveLength(1)
    expect(history.changes[0]).toMatchObject({
      action: 'Data.SetField',
      source: 'link',
      docId: DOC,
    })
    expect(history.changes[0]!.changes.map((c) => c.objectId)).toEqual(['f_name', 'f_qty', 'f_new'])
    expect(await zip.file('[Content_Types].xml')!.async('string')).toContain('itemProps')
    // a second write appends to the same history part
    const again = await writeFieldsToXlsx(bytes, {
      docId: DOC,
      writes: [write('f_qty', 8, { rev: 1, text: '7' })],
    })
    expect(
      parseHistoryXml((await readCustomXmlPart(again.bytes!, HISTORY_NS))!.xml).changes,
    ).toHaveLength(2)
  })

  it('reports conflicts (DVH revision or an Excel edit), formulas and unknown fields', async () => {
    // Excel changed the name cell: the model part still says rev 2 / "Dự án A"
    const out = await writeFieldsToXlsx(await workbook('Sửa trong Excel'), {
      docId: DOC,
      writes: [
        write('f_name', 'Dự án B', { rev: 2, text: 'Dự án A' }),
        write('f_qty', 9, { rev: 1, text: '5' }),
        write('f_sum', 99, null),
        write('f_gone', 1, null),
      ],
    })
    expect(out.results.map((r) => r.status)).toEqual(['conflict', 'conflict', 'formula', 'missing'])
    expect(out.results[0]!.current).toEqual({ rev: 2, text: 'Sửa trong Excel' })
    expect(out.bytes).toBeNull()
    // keep this side: forced
    const forced = await writeFieldsToXlsx(await workbook('Sửa trong Excel'), {
      docId: DOC,
      writes: [write('f_name', 'Dự án B', { rev: 2, text: 'Dự án A' }, true)],
    })
    expect(forced.results[0]!.status).toBe('written')
  })

  it('refuses another workbook', async () => {
    await expect(
      writeFieldsToXlsx(await workbook(), {
        docId: 'doc_other',
        writes: [write('f_qty', 1, null)],
      }),
    ).rejects.toThrow()
  })

  it('setCellInSheetXml inserts into self-closed rows and empty sheets', () => {
    expect(setCellInSheetXml('<sheetData/>', 0, 1, 3)).toBe(
      '<sheetData><row r="1"><c r="B1"><v>3</v></c></row></sheetData>',
    )
    expect(setCellInSheetXml('<sheetData><row r="2"/></sheetData>', 1, 0, true)).toBe(
      '<sheetData><row r="2"><c r="A2" t="b"><v>1</v></c></row></sheetData>',
    )
    expect(
      setCellInSheetXml(
        '<sheetData><row r="1"><c r="A1"/><c r="C1"/></row></sheetData>',
        0,
        1,
        'x',
      ),
    ).toBe(
      '<sheetData><row r="1"><c r="A1"/><c r="B1" t="inlineStr"><is><t xml:space="preserve">x</t></is></c><c r="C1"/></row></sheetData>',
    )
  })
})

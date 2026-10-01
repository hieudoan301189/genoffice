// Spike S3 (docs/dvh-architecture-implementation-plan.md): DVH binds xlsx cells
// with hidden defined names (`_dvh.f.<id>`, `_dvh.t.<id>`). The editor never
// models hidden names, so every structural save must carry them along.
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyDefinedNamesState } from '../src/gateway/xlsx-defined-names'
import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import type { SheetEditPlan } from '../src/gateway/xlsx-sheets'
import type { StructuralOp } from '../src/gateway/xlsx-structure'
import type { DvhWorkbookState } from '../src/gateway/xlsx-dvh'

const FIELD = '_dvh.f.f_projectname0001'
const TABLE = '_dvh.t.t_workitems000001'
const MODEL_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<dvh:model xmlns:dvh="urn:dvh-office:model:1" docId="doc_test"><dvh:fields/></dvh:model>'

async function workbook(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>' +
      `<definedNames><definedName name="${FIELD}" hidden="1">Data!$B$5</definedName>` +
      `<definedName name="${TABLE}" hidden="1">Data!$A$8:$C$11</definedName></definedNames>` +
      '</workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>',
  )
  const sheet = (rows: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows}</sheetData></worksheet>`
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet(
      '<row r="5"><c r="A5" t="inlineStr"><is><t>Tên dự án</t></is></c><c r="B5" t="inlineStr"><is><t>Dự án A</t></is></c></row>' +
        '<row r="8"><c r="A8" t="inlineStr"><is><t>Mã</t></is></c></row>',
    ),
  )
  zip.file('xl/worksheets/sheet2.xml', sheet(''))
  zip.file('customXml/item1.xml', MODEL_XML)
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function namesAfter(
  ops: readonly StructuralOp[],
  sheetPlan?: SheetEditPlan,
  dvhState: DvhWorkbookState | null = null,
): Promise<{ names: Map<string, { formula: string; hidden: boolean }>; zip: JSZip }> {
  const structural = ops.length > 0 ? [{ sheetName: 'Data', ops }] : []
  const mutation = await applyCellEditsToXlsx(
    await workbook(),
    [],
    structural,
    [],
    sheetPlan,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    dvhState,
  )
  const zip = await JSZip.loadAsync(mutation.buffer)
  const xml = await zip.file('xl/workbook.xml')!.async('string')
  const names = new Map<string, { formula: string; hidden: boolean }>()
  for (const m of xml.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)) {
    names.set(/name="([^"]*)"/.exec(m[1]!)![1]!, {
      formula: m[2]!,
      hidden: /hidden="(?:1|true)"/.test(m[1]!),
    })
  }
  return { names, zip }
}

describe('hidden DVH names across structural saves', () => {
  it.each<[string, StructuralOp[], string, string]>([
    [
      'insert rows above',
      [{ kind: 'insert-rows', index: 2, count: 1 }],
      'Data!$B$6',
      'Data!$A$9:$C$12',
    ],
    [
      'insert a row inside the table',
      [{ kind: 'insert-rows', index: 9, count: 1 }],
      'Data!$B$5',
      'Data!$A$8:$C$12',
    ],
    [
      'remove rows above',
      [{ kind: 'remove-rows', index: 1, count: 2 }],
      'Data!$B$3',
      'Data!$A$6:$C$9',
    ],
    [
      'insert a column before B',
      [{ kind: 'insert-cols', index: 1, count: 1 }],
      'Data!$C$5',
      'Data!$A$8:$D$11',
    ],
    [
      'move row 5 to the top',
      [{ kind: 'move-rows', index: 4, count: 1, before: 0 }],
      'Data!$B$1',
      'Data!$A$8:$C$11',
    ],
  ])('%s: both names follow and stay hidden', async (_label, ops, field, table) => {
    const { names, zip } = await namesAfter(ops)
    expect(names.get(FIELD)).toEqual({ formula: field, hidden: true })
    expect(names.get(TABLE)).toEqual({ formula: table, hidden: true })
    expect(await zip.file('customXml/item1.xml')!.async('string')).toBe(MODEL_XML)
  })

  // S3 finding: the editor never sees hidden names, so the user deletes freely
  // and only the save fails. P1 must resolve or unbind the binding first.
  it('removing the bound row currently aborts the whole save', async () => {
    await expect(namesAfter([{ kind: 'remove-rows', index: 4, count: 1 }])).rejects.toThrow(
      /references the deleted range \(\$B\$5\)/,
    )
  })

  it('removing rows that only clip the table range still saves', async () => {
    const { names } = await namesAfter([{ kind: 'remove-rows', index: 10, count: 1 }])
    expect(names.get(FIELD)).toEqual({ formula: 'Data!$B$5', hidden: true })
    expect(names.get(TABLE)).toEqual({ formula: 'Data!$A$8:$C$10', hidden: true })
  })

  it('a sheet rename rewrites the hidden names', async () => {
    const { names } = await namesAfter([], {
      renames: [{ sheetName: 'Data', newName: 'Dữ liệu' }],
      additions: [],
      removals: [],
      order: ['Dữ liệu', 'Other'],
    })
    expect(names.get(FIELD)).toEqual({ formula: "'Dữ liệu'!$B$5", hidden: true })
  })

  // S3 finding: same for sheets — a binding makes its sheet undeletable at save time.
  it('a hidden DVH name currently blocks deleting the sheet it binds', async () => {
    await expect(
      namesAfter([], { renames: [], additions: [], removals: ['Data'], order: ['Other'] }),
    ).rejects.toThrow(/defined name references "Data"/)
  })

  it('the Name Manager save keeps hidden names verbatim', async () => {
    const zip = await JSZip.loadAsync(await workbook())
    const xml = await zip.file('xl/workbook.xml')!.async('string')
    const saved = applyDefinedNamesState(xml, {
      names: [{ name: 'TenDuAn', formula: 'Data!$B$5' }],
      preserveNames: [],
    })
    expect(saved).toContain(`<definedName name="${FIELD}" hidden="1">Data!$B$5</definedName>`)
    expect(saved).toContain('<definedName name="TenDuAn">Data!$B$5</definedName>')
  })

  // P1: the editor models the binding names (system names) and sends them on save
  describe('with the editor-modeled names (P1)', () => {
    const state = (field: string, table = 'Data!$A$8:$C$11'): DvhWorkbookState => ({
      names: [
        { name: FIELD, formula: field },
        { name: TABLE, formula: table },
      ],
      customXmlParts: [],
    })

    it("deleting the bound row saves, writing the editor's #REF! name", async () => {
      const { names } = await namesAfter(
        [{ kind: 'remove-rows', index: 4, count: 1 }],
        undefined,
        state('Data!#REF!', 'Data!$A$7:$C$10'),
      )
      expect(names.get(FIELD)).toEqual({ formula: 'Data!#REF!', hidden: true })
      expect(names.get(TABLE)).toEqual({ formula: 'Data!$A$7:$C$10', hidden: true })
    })

    it('a cut/paste move is written where the editor tracked it', async () => {
      const { names } = await namesAfter([], undefined, state('Data!$E$6'))
      expect(names.get(FIELD)).toEqual({ formula: 'Data!$E$6', hidden: true })
    })

    it('the editor names are not shifted a second time by structural ops', async () => {
      const { names } = await namesAfter(
        [{ kind: 'insert-rows', index: 2, count: 1 }],
        undefined,
        state('Data!$B$6', 'Data!$A$9:$C$12'),
      )
      expect(names.get(FIELD)).toEqual({ formula: 'Data!$B$6', hidden: true })
    })

    it('a bound sheet can be deleted once the editor dropped its bindings', async () => {
      const { names } = await namesAfter(
        [],
        { renames: [], additions: [], removals: ['Data'], order: ['Other'] },
        { names: [], customXmlParts: [] },
      )
      expect(names.has(FIELD)).toBe(false)
    })

    it('writes the model part in place or adds it with relationships', async () => {
      const model = (v: string) =>
        `<dvh:model xmlns:dvh="urn:dvh-office:model:1" docId="doc_test"><dvh:fields><dvh:f id="f_a">${v}</dvh:f></dvh:fields></dvh:model>`
      const replaced = await namesAfter([], undefined, {
        names: [],
        customXmlParts: [{ ns: 'urn:dvh-office:model:1', xml: model('A'), storeItemId: '{X}' }],
      })
      expect(await replaced.zip.file('customXml/item1.xml')!.async('string')).toContain('>A<')
      const added = await namesAfter([], undefined, {
        names: [],
        customXmlParts: [
          {
            ns: 'urn:dvh-office:history:1',
            xml: '<dvh:history xmlns:dvh="urn:dvh-office:history:1"/>',
            storeItemId: '{H}',
          },
        ],
      })
      const zip = added.zip
      expect(await zip.file('customXml/item2.xml')!.async('string')).toContain(
        'urn:dvh-office:history:1',
      )
      expect(await zip.file('customXml/itemProps2.xml')!.async('string')).toContain(
        'ds:itemID="{H}"',
      )
      expect(await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).toContain(
        'Target="../customXml/item2.xml"',
      )
      expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
        'PartName="/customXml/itemProps2.xml"',
      )
    })
  })
})

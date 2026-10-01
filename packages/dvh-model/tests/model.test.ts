import { describe, expect, it } from 'vitest'
import {
  diffModels,
  emptyModel,
  fieldDefinedName,
  fieldSdtPrXml,
  fieldSdtTag,
  fieldXPath,
  findPartByNamespace,
  HISTORY_NS,
  isDvhDefinedName,
  isDvhId,
  itemPropsXml,
  MODEL_NS,
  modelHash,
  newDvhId,
  newStoreItemId,
  parseDvhDefinedName,
  parseDvhSdtTag,
  parseHistoryXml,
  parseModelXml,
  serializeHistoryXml,
  serializeModelXml,
  setFieldTextInXml,
  storeItemIdOf,
  type DvhModel,
} from '../src/index'

// ADR D10: the core must work with any schema. Two unrelated vocabularies:
const acceptance: DvhModel = {
  docId: 'doc_aaaaaaaaaaaaaaaa',
  schemaVersion: 1,
  fields: [
    {
      id: 'f_projectname00001',
      name: 'Project.Name',
      type: 'text',
      value: 'Dự án Cầu Bến Thủy 3 & <nhánh>',
      access: 'readwrite',
    },
    {
      id: 'f_acceptdate000001',
      name: 'Acceptance.Date',
      type: 'date',
      value: '01/10/2026',
      access: 'readwrite',
    },
    {
      id: 'f_passed0000000001',
      name: 'Acceptance.Passed',
      type: 'boolean',
      value: true,
      access: 'read',
    },
  ],
  collections: [
    {
      id: 'c_workitems0000001',
      name: 'WorkItems',
      columns: [
        { id: 'col_code', key: 'code', title: 'Mã', type: 'text' },
        { id: 'col_qty', key: 'qty', title: 'Khối lượng', type: 'number', numFmt: '#,##0.000' },
      ],
      rows: [
        ['CV01', 125.5],
        ['CV02', null],
      ],
    },
  ],
  links: [],
}
const inventory: DvhModel = {
  docId: 'doc_bbbbbbbbbbbbbbbb',
  schemaVersion: 1,
  fields: [
    {
      id: 'f_warehouse0000001',
      name: 'Kho.Ten',
      type: 'enum',
      value: 'Kho B',
      enumValues: ['Kho A', 'Kho B'],
      access: 'readwrite',
    },
    {
      id: 'f_totalqty00000001',
      name: 'Kho.TongSoLuong',
      type: 'number',
      value: 1234.5,
      access: 'read',
    },
    {
      id: 'f_note000000000001',
      name: 'Kho.GhiChu',
      type: 'text',
      value: null,
      access: 'readwrite',
    },
  ],
  collections: [],
  links: [
    {
      id: 'l_fromsheet0000001',
      source: {
        docId: 'doc_aaaaaaaaaaaaaaaa',
        relPath: '../QLCL/data.xlsx',
        objectId: 'f_projectname00001',
      },
      targets: ['f_note000000000001'],
      update: 'manual',
    },
  ],
}

describe('ids', () => {
  it('mints well-formed, distinct ids and store item GUIDs', () => {
    const ids = new Set(Array.from({ length: 200 }, () => newDvhId('f')))
    expect(ids.size).toBe(200)
    for (const id of ids) expect(isDvhId(id, 'f')).toBe(true)
    expect(isDvhId(newDvhId('doc'), 'f')).toBe(false)
    expect(isDvhId('f_short')).toBe(false)
    expect(newStoreItemId()).toMatch(
      /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/,
    )
  })
})

describe('model part', () => {
  it.each([
    ['acceptance', acceptance],
    ['inventory', inventory],
  ])('round-trips the %s schema', (_name, model) => {
    const xml = serializeModelXml(model)
    expect(xml).toContain(`xmlns:dvh="${MODEL_NS}"`)
    expect(parseModelXml(xml)).toEqual(model)
  })

  it('reads a part Word re-serialized (attribute order, escaped JSON instead of CDATA, no declaration)', () => {
    const xml =
      `<dvh:model schemaVersion="1" docId="doc_aaaaaaaaaaaaaaaa" xmlns:dvh="${MODEL_NS}"><dvh:fields>` +
      '<dvh:f type="text" access="readwrite" name="Project.Name" id="f_projectname00001">Sửa trong Word</dvh:f>' +
      '</dvh:fields><dvh:objects>{"collections":[],"links":[]}</dvh:objects></dvh:model>'
    const model = parseModelXml(xml)
    expect(model.fields[0]).toEqual({
      id: 'f_projectname00001',
      name: 'Project.Name',
      type: 'text',
      access: 'readwrite',
      value: 'Sửa trong Word',
    })
  })

  it('rejects parts that are not a DVH model', () => {
    expect(() => parseModelXml('<root xmlns="urn:other"/>')).toThrow()
  })

  it('updates one field in place without touching the rest of the part', () => {
    const xml = serializeModelXml(acceptance)
    const updated = setFieldTextInXml(xml, 'f_projectname00001', 'Dự án mới <A>')!
    expect(parseModelXml(updated).fields[0]!.value).toBe('Dự án mới <A>')
    expect(updated.replace(/<dvh:f id="f_projectname00001"[^]*?<\/dvh:f>/, '')).toBe(
      xml.replace(/<dvh:f id="f_projectname00001"[^]*?<\/dvh:f>/, ''),
    )
    expect(setFieldTextInXml(xml, 'f_missing', 'x')).toBeNull()
  })

  it('finds parts by namespace whatever their number', () => {
    const parts = [
      { path: 'customXml/item1.xml', xml: serializeHistoryXml({ docId: 'doc_x', changes: [] }) },
      { path: 'customXml/itemProps1.xml', xml: itemPropsXml('{A}', HISTORY_NS) },
      { path: 'customXml/item2.xml', xml: serializeModelXml(inventory) },
    ]
    expect(findPartByNamespace(parts, MODEL_NS)?.path).toBe('customXml/item2.xml')
    expect(findPartByNamespace(parts, HISTORY_NS)?.path).toBe('customXml/item1.xml')
    expect(findPartByNamespace(parts, 'urn:none')).toBeNull()
    expect(storeItemIdOf(itemPropsXml('{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A01}', MODEL_NS))).toBe(
      '{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A01}',
    )
  })

  it('builds the binding XPath Word resolves', () => {
    expect(fieldXPath('f_projectname00001')).toBe(
      "/dvh:model[1]/dvh:fields[1]/dvh:f[@id='f_projectname00001'][1]",
    )
  })
})

describe('history part', () => {
  it('round-trips change sets and the model hash', () => {
    const changes = [
      {
        id: 'cs_1',
        txId: 'tx_1',
        docId: acceptance.docId,
        at: '2026-10-02T08:00:00.000Z',
        source: 'ui' as const,
        action: 'Data.SetField',
        changes: [{ objectId: 'f_projectname00001', path: 'value', before: 'A', after: 'B ]]> C' }],
      },
    ]
    const hash = modelHash(acceptance)
    const parsed = parseHistoryXml(
      serializeHistoryXml({ docId: acceptance.docId, modelHash: hash, changes }),
    )
    expect(parsed).toEqual({ docId: acceptance.docId, modelHash: hash, changes })
  })

  it('hashes models independent of key order and sensitive to values', () => {
    const reverseKeys = (value: unknown): unknown =>
      Array.isArray(value)
        ? value.map(reverseKeys)
        : value && typeof value === 'object'
          ? Object.fromEntries(
              Object.entries(value)
                .reverse()
                .map(([k, v]) => [k, reverseKeys(v)]),
            )
          : value
    const reordered = reverseKeys(acceptance) as DvhModel
    expect(Object.keys(reordered)[0]).toBe('links')
    expect(modelHash(reordered)).toBe(modelHash(acceptance))
    const edited = {
      ...acceptance,
      fields: acceptance.fields.map((f, i) => (i === 0 ? { ...f, value: 'x' } : f)),
    }
    expect(modelHash(edited)).not.toBe(modelHash(acceptance))
  })

  it('diffs field values, renames, additions and removals', () => {
    const after: DvhModel = {
      ...acceptance,
      fields: [
        { ...acceptance.fields[0]!, value: 'Dự án mới' },
        { ...acceptance.fields[1]!, name: 'Acceptance.Day' },
        { id: 'f_new0000000000001', name: 'New', type: 'text', value: null, access: 'readwrite' },
      ],
    }
    const changes = diffModels(acceptance, after)
    expect(changes.map((c) => `${c.objectId}:${c.path}`)).toEqual([
      'f_projectname00001:value',
      'f_acceptdate000001:name',
      'f_new0000000000001:',
      'f_passed0000000001:',
    ])
    expect(diffModels(emptyModel('doc_x'), emptyModel('doc_x'))).toEqual([])
  })
})

describe('binding conventions', () => {
  it('names and parses tags and defined names', () => {
    expect(fieldSdtTag('f_a')).toBe('dvh:f:f_a')
    expect(fieldDefinedName('f_a')).toBe('_dvh.f.f_a')
    expect(parseDvhDefinedName('_dvh.t.t_b')).toEqual({ kind: 'table', id: 't_b' })
    expect(parseDvhSdtTag('dvh:f:f_a')).toEqual({ kind: 'field', id: 'f_a' })
    expect(isDvhDefinedName('_dvh.f.f_a')).toBe(true)
    expect(isDvhDefinedName('TenDuAn')).toBe(false)
    expect(parseDvhDefinedName('TenDuAn')).toBeNull()
  })

  it('builds the content-control properties Word binds (attribute values escaped)', () => {
    const xml = fieldSdtPrXml({
      fieldId: 'f_a',
      alias: 'Tên "dự án"',
      sdtId: 42.7,
      storeItemId: '{ID}',
    })
    expect(xml).toBe(
      '<w:sdtPr><w:alias w:val="Tên &quot;dự án&quot;"/><w:tag w:val="dvh:f:f_a"/><w:id w:val="42"/>' +
        `<w:dataBinding w:prefixMappings="xmlns:dvh='urn:dvh-office:model:1'" w:xpath="/dvh:model[1]/dvh:fields[1]/dvh:f[@id='f_a'][1]" w:storeItemID="{ID}"/>` +
        '<w:text/></w:sdtPr>',
    )
  })
})

// Spike S1 (docs/dvh-architecture-implementation-plan.md): the OOXML carriers for
// the DVH model, shared by make-fixtures.ts and verify.ts. Throwaway spike code:
// the production version belongs in packages/dvh-model once P0 freezes the spec.

export const MODEL_NS = 'urn:dvh-office:model:1'
export const HISTORY_NS = 'urn:dvh-office:history:1'
export const CUSTOM_XML_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
export const CUSTOM_XML_PROPS_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
export const CUSTOM_XML_PROPS_TYPE =
  'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'

export const DOCX_DOC_ID = 'doc_s1docx00000001'
export const XLSX_DOC_ID = 'doc_s1xlsx00000001'
export const FIELD_PROJECT = 'f_projectname0001'
export const FIELD_DATE = 'f_acceptdate00001'
export const TABLE_ITEMS = 't_workitems000001'

export const PROJECT_VALUE = 'Dự án Cầu Bến Thủy 3'
export const DATE_VALUE = '01/10/2026'
/** what the bound content controls show before Word refreshes them from the store */
export const STALE_A = 'STALE-A'
export const STALE_B = 'STALE-B'
/** text Word writes into the project content control (two-way binding probe) */
export const WORD_EDIT_VALUE = 'Sửa trong Word – Gói thầu số 7'

/** store item ids: Word binds content controls to a part through this GUID */
export const MODEL_STORE_ID = '{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A01}'
export const HISTORY_STORE_ID = '{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A02}'

/** XPath forms under test: attribute predicate (A) and positional (B) */
export const XPATH_BY_ID = `/dvh:model[1]/dvh:fields[1]/dvh:f[@id='${FIELD_PROJECT}'][1]`
export const XPATH_BY_POSITION = '/dvh:model[1]/dvh:fields[1]/dvh:f[2]'
export const PREFIX_MAPPINGS = `xmlns:dvh='${MODEL_NS}'`

export const DEFINED_FIELD = `_dvh.f.${FIELD_PROJECT}`
export const DEFINED_TABLE = `_dvh.t.${TABLE_ITEMS}`

export const WORK_ITEMS: [string, string, number][] = [
  ['CV01', 'Đào đất hố móng', 125.5],
  ['CV02', 'Bê tông lót móng M100', 18.25],
  ['CV03', 'Cốt thép móng D<=18', 2.734],
]

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

export function modelPartXml(docId: string): string {
  const objects = {
    tables: [
      {
        id: TABLE_ITEMS,
        name: 'WorkItems',
        columns: [
          { id: 'col_code', title: 'Mã', type: 'text' },
          { id: 'col_name', title: 'Công việc', type: 'text' },
          { id: 'col_qty', title: 'Khối lượng', type: 'number', numFmt: '#,##0.000' },
        ],
        rows: WORK_ITEMS,
      },
    ],
    links: [],
  }
  return (
    XML_DECL +
    `<dvh:model xmlns:dvh="${MODEL_NS}" docId="${docId}" schemaVersion="1">` +
    '<dvh:fields>' +
    `<dvh:f id="${FIELD_PROJECT}" name="Project.Name" type="text">${escapeXml(PROJECT_VALUE)}</dvh:f>` +
    `<dvh:f id="${FIELD_DATE}" name="Acceptance.Date" type="date">${escapeXml(DATE_VALUE)}</dvh:f>` +
    '</dvh:fields>' +
    `<dvh:objects><![CDATA[${JSON.stringify(objects)}]]></dvh:objects>` +
    '</dvh:model>'
  )
}

/** change-set history; `targetBytes` pads it with realistic entries for the size probe */
export function historyPartXml(docId: string, targetBytes: number): string {
  const lines: string[] = []
  let size = 0
  let seq = 0
  while (size < targetBytes || lines.length === 0) {
    const entry = {
      id: `cs_${seq.toString(36).padStart(8, '0')}`,
      txId: `tx_${Math.floor(seq / 3).toString(36)}`,
      docId,
      at: new Date(Date.UTC(2026, 8, 1) + seq * 61_000).toISOString(),
      source: ['ui', 'ai', 'workflow', 'link'][seq % 4],
      action: seq % 2 ? 'Data.SetField' : 'Table.Refresh',
      changes: [
        {
          objectId: seq % 2 ? FIELD_PROJECT : TABLE_ITEMS,
          path: seq % 2 ? 'value' : `rows[${seq % 40}][2]`,
          before: `${PROJECT_VALUE} – bản ${seq}`,
          after: `${PROJECT_VALUE} – bản ${seq + 1} (${(seq * 7919) % 100_003})`,
        },
      ],
    }
    const line = JSON.stringify(entry)
    lines.push(line)
    size += line.length + 1
    seq++
  }
  return (
    XML_DECL +
    `<dvh:history xmlns:dvh="${HISTORY_NS}" docId="${docId}" count="${lines.length}">` +
    `<dvh:changes><![CDATA[${lines.join('\n')}]]></dvh:changes>` +
    '</dvh:history>'
  )
}

export function itemPropsXml(storeId: string, ns: string): string {
  return (
    XML_DECL +
    `<ds:datastoreItem ds:itemID="${storeId}" ` +
    'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
    `<ds:schemaRefs><ds:schemaRef ds:uri="${ns}"/></ds:schemaRefs></ds:datastoreItem>`
  )
}

export function itemRelsXml(propsTarget: string): string {
  return (
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="${CUSTOM_XML_PROPS_REL}" Target="${propsTarget}"/>` +
    '</Relationships>'
  )
}

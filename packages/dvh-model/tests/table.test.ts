import { describe, expect, it } from 'vitest'
import {
  collectionDefinedName,
  emptyModel,
  formatCellValue,
  isDvhDefinedName,
  parseDvhDefinedName,
  parseModelXml,
  renderTable,
  serializeModelXml,
  tableSdtPrXml,
  type DvhCollection,
  type DvhTable,
} from '../src/index'

// two unrelated schemas (ADR D10): work items of an acceptance file, a materials list
const workItems: DvhCollection = {
  id: 'c_workitems0000001',
  name: 'WorkItems',
  columns: [
    {
      id: 'k_code',
      key: 'code',
      title: 'Mã',
      type: 'text',
      headerStyle: { bold: true, fill: '#FFF2CC' },
      style: { align: 'center' },
      width: 48,
    },
    { id: 'k_name', key: 'name', title: 'Công việc', type: 'text', width: 200 },
    {
      id: 'k_qty',
      key: 'qty',
      title: 'Khối lượng',
      type: 'number',
      numFmt: '#,##0.00',
      style: { color: '#C00000' },
    },
  ],
  rows: [
    ['AB.1', 'Đào móng', 120.5],
    ['AB.2', 'Bê tông lót', 18],
    ['AB.3', 'Cốt thép', 2450.25],
  ],
}

const materials: DvhCollection = {
  id: 'c_materials000001',
  name: 'Materials',
  columns: [
    { id: 'k_mat', key: 'material', title: 'Vật tư', type: 'text' },
    { id: 'k_unit', key: 'unit', title: 'ĐVT', type: 'text' },
    { id: 'k_price', key: 'price', title: 'Đơn giá', type: 'number', numFmt: '#,##0 "đ"' },
  ],
  rows: [
    ['Xi măng', 'tấn', 1650000],
    ['Thép D10', 'kg', 18500],
  ],
}

const tableOf = (collection: DvhCollection, mode: 'source' | 'destination'): DvhTable => ({
  id: `t_${collection.id.slice(2)}`,
  name: collection.name,
  source: { collectionId: collection.id },
  columns: collection.columns.map((c) => ({ id: `tc_${c.id}`, columnId: c.id, title: c.title })),
  style: { mode },
  layout: { repeatHeader: true, keepRowsTogether: true, widths: 'page' },
})

describe('DVH.Table model', () => {
  it('round-trips tables through the model part and reads P1 files without them', () => {
    const model = {
      ...emptyModel('doc_aaaaaaaaaaaaaaaa'),
      collections: [workItems, materials],
      tables: [tableOf(workItems, 'source'), tableOf(materials, 'destination')],
    }
    const back = parseModelXml(serializeModelXml(model))
    expect(back.tables).toEqual(model.tables)
    expect(back.collections[0]?.columns[0]?.headerStyle).toEqual({ bold: true, fill: '#FFF2CC' })
    const p1 = serializeModelXml(emptyModel('doc_bbbbbbbbbbbbbbbb')).replace(',"tables":[]', '')
    expect(parseModelXml(p1).tables).toEqual([])
  })

  it('names collection ranges as system names and tags table controls', () => {
    const name = collectionDefinedName(workItems.id)
    expect(name).toBe('_dvh.c.c_workitems0000001')
    expect(isDvhDefinedName(name)).toBe(true)
    expect(parseDvhDefinedName(name)).toEqual({ kind: 'collection', id: workItems.id })
    expect(tableSdtPrXml({ tableId: 't_x', alias: 'Bảng <1>', sdtId: 7 })).toBe(
      '<w:sdtPr><w:alias w:val="Bảng &lt;1&gt;"/><w:tag w:val="dvh:t:t_x"/><w:id w:val="7"/></w:sdtPr>',
    )
  })
})

describe('renderTable', () => {
  it('keeps the source formatting in source mode', () => {
    const grid = renderTable(tableOf(workItems, 'source'), [workItems])
    expect(grid.headerRowCount).toBe(1)
    expect(grid.bodyRowCount).toBe(3)
    expect(grid.rows[0]!.map((c) => c.value)).toEqual(['Mã', 'Công việc', 'Khối lượng'])
    expect(grid.rows[0]![0]!.style).toEqual({ bold: true, fill: '#FFF2CC' })
    expect(grid.rows[1]![0]!.style).toEqual({ align: 'center' })
    expect(grid.rows[3]![2]).toMatchObject({
      value: 2450.25,
      numFmt: '#,##0.00',
      style: { color: '#C00000' },
    })
    // 48 / 200 / default 64 px
    expect(grid.widths.map((w) => Math.round(w * 312))).toEqual([48, 200, 64])
  })

  it('applies the destination style with defaults, bands and number alignment', () => {
    const table = { ...tableOf(materials, 'destination') }
    table.style = { mode: 'destination', header: { fill: '#E2EFDA' }, bandFill: '#F2F2F2' }
    const grid = renderTable(table, [materials])
    expect(grid.rows[0]![0]!.style).toMatchObject({ bold: true, fill: '#E2EFDA', align: 'center' })
    expect(grid.rows[1]![2]!.style.align).toBe('right')
    expect(grid.rows[1]![0]!.style.fill).toBeUndefined()
    expect(grid.rows[2]![0]!.style.fill).toBe('#F2F2F2')
    expect(grid.rows[2]![2]!.style.border).toEqual({ color: '#7F7F7F', width: 'thin' })
  })

  it('adds group header rows and a total row', () => {
    const table: DvhTable = {
      ...tableOf(workItems, 'destination'),
      headerGroups: [
        [
          { title: 'Hạng mục', span: 2 },
          { title: 'Số liệu', span: 1 },
        ],
      ],
      totalRow: { label: 'Tổng' },
    }
    table.columns[2] = { ...table.columns[2]!, total: 'sum' }
    const grid = renderTable(table, [workItems])
    expect(grid.headerRowCount).toBe(2)
    expect(grid.rows[0]!.map((c) => [c.value, c.colSpan])).toEqual([
      ['Hạng mục', 2],
      ['Số liệu', 1],
    ])
    const totals = grid.rows[grid.rows.length - 1]!
    expect(totals.map((c) => c.value)).toEqual(['Tổng', null, 2588.75])
    expect(totals[2]).toMatchObject({ kind: 'total', numFmt: '#,##0.00', style: { bold: true } })
    expect(() =>
      renderTable({ ...table, headerGroups: [[{ title: 'x', span: 2 }]] }, [workItems]),
    ).toThrow('spans 2 of 3')
  })

  it('follows the collection when rows are added or removed, and embedded rows work alone', () => {
    const grown = { ...workItems, rows: [...workItems.rows, ['AB.4', 'Ván khuôn', 75]] }
    expect(renderTable(tableOf(workItems, 'source'), [grown]).bodyRowCount).toBe(4)
    const embedded: DvhTable = {
      ...tableOf(materials, 'destination'),
      source: { rows: [['Cát', 'm3', 350000]] },
    }
    expect(renderTable(embedded, []).rows[1]!.map((c) => c.value)).toEqual(['Cát', 'm3', 350000])
    expect(() => renderTable(tableOf(workItems, 'source'), [])).toThrow('not found')
  })
})

describe('formatCellValue', () => {
  it('formats digits, grouping, decimals, percent, literal text and dates', () => {
    expect(formatCellValue(1250000, '#,##0', 'en-US')).toBe('1,250,000')
    expect(formatCellValue(1250000, '#,##0', 'vi-VN')).toBe('1.250.000')
    expect(formatCellValue(2450.25, '#,##0.00', 'vi-VN')).toBe('2.450,25')
    expect(formatCellValue(1650000, '#,##0 "đ"', 'vi-VN')).toBe('1.650.000 đ')
    expect(formatCellValue(0.125, '0.0%', 'en-US')).toBe('12.5%')
    expect(formatCellValue(18, '0.00', 'en-US')).toBe('18.00')
    expect(formatCellValue(45931, 'dd/mm/yyyy')).toBe('01/10/2025')
    expect(formatCellValue(0.1 + 0.2, undefined, 'en-US')).toBe('0.3')
    expect(formatCellValue(null, '#,##0')).toBe('')
    expect(formatCellValue('AB.1', '#,##0')).toBe('AB.1')
  })
})

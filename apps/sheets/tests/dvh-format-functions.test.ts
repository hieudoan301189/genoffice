import { describe, expect, it } from 'vitest'
import { workbookCellEditSchema } from '../src/shared/desktop-api'
import {
  FormulaFormatStore,
  composeFormattedStyle,
  withBakedFormulaFormats,
  type FormatOwner,
} from '../src/renderer/dvh-format-channel'
import {
  a1Address,
  copiedCellStyle,
  isFalseLike,
  oleColorToHex,
  parseDvhFontCodes,
  parseTitleRows,
  selectTableRows,
} from '../src/renderer/dvh-format-functions'

const owner = (row: number, column: number, functionName = 'DVH.FONT'): FormatOwner => ({
  unitId: 'u1',
  sheetId: 's1',
  row,
  column,
  functionName,
})
const point = (row: number, column: number, style: object) => ({
  sheetId: 's1',
  startRow: row,
  endRow: row,
  startColumn: column,
  endColumn: column,
  style,
})

describe('FormulaFormatStore', () => {
  it('serves area and point formats, later publishes winning on overlap', () => {
    const store = new FormulaFormatStore()
    store.replace(owner(0, 5), [
      {
        sheetId: 's1',
        startRow: 0,
        endRow: 9,
        startColumn: 0,
        endColumn: 1,
        style: { bold: true, fillColor: '#FF0000' },
      },
    ])
    store.replace(owner(0, 6, 'DVH.FILLCOLOR'), [point(3, 1, { fillColor: '#00FF00' })])
    expect(store.styleAt('u1', 's1', 0, 0)).toEqual({ bold: true, fillColor: '#FF0000' })
    expect(store.styleAt('u1', 's1', 3, 1)).toEqual({ bold: true, fillColor: '#00FF00' })
    expect(store.styleAt('u1', 's1', 10, 0)).toBeUndefined()
    expect(store.styleAt('u1', 's2', 0, 0)).toBeUndefined()
    expect(store.styleAt('u2', 's1', 0, 0)).toBeUndefined()
  })

  it('reports no change (no repaint) when a recalc republishes identical formats', () => {
    const store = new FormulaFormatStore()
    const events: string[] = []
    store.onChange((unitId) => events.push(unitId))
    expect(store.replace(owner(0, 0), [point(1, 1, { italic: true })])).toBe(true)
    expect(store.replace(owner(0, 0), [point(1, 1, { italic: true })])).toBe(false)
    expect(store.replace(owner(0, 0), [point(2, 1, { italic: true })])).toBe(true)
    expect(store.styleAt('u1', 's1', 1, 1)).toBeUndefined()
    expect(store.styleAt('u1', 's1', 2, 1)).toEqual({ italic: true })
    expect(events).toEqual(['u1', 'u1'])
  })

  it('withdraws formats when the owner clears, publishes nothing, or its unit goes away', () => {
    const store = new FormulaFormatStore()
    store.replace(owner(0, 0), [point(1, 1, { bold: true })])
    store.replace(owner(0, 0), [])
    expect(store.styleAt('u1', 's1', 1, 1)).toBeUndefined()
    store.replace(owner(0, 0), [point(1, 1, { bold: true })])
    store.clear(owner(0, 0))
    expect(store.ownersOf('u1')).toEqual([])
    store.replace(owner(0, 0), [point(1, 1, { bold: true })])
    store.clearUnit('u1')
    expect(store.styleAt('u1', 's1', 1, 1)).toBeUndefined()
  })

  it('knows which format-only edits must recalculate a format-reading owner', () => {
    const store = new FormulaFormatStore()
    const table = owner(0, 6, 'DVH.TABLE')
    store.watchSources(table, [
      { sheetId: 's1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 3 },
    ])
    const range = (startRow: number, startColumn: number) => ({
      startRow,
      endRow: startRow,
      startColumn,
      endColumn: startColumn,
    })
    expect(store.isWatched('u1', 's1', range(1, 1))).toBe(true)
    expect(store.isWatched('u1', 's1', range(5, 1))).toBe(false)
    expect(store.isWatched('u1', 's2', range(1, 1))).toBe(false)
    store.clear(table)
    expect(store.isWatched('u1', 's1', range(1, 1))).toBe(false)
  })

  it('bakes every formatted cell once with its merged patch', () => {
    const store = new FormulaFormatStore()
    store.replace(owner(9, 9), [
      {
        sheetId: 's1',
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 0,
        style: { bold: true },
      },
    ])
    store.replace(owner(9, 8), [point(1, 0, { fontColor: '#0000FF' })])
    expect(store.bake('u1')).toEqual(
      expect.arrayContaining([
        { sheetId: 's1', row: 0, column: 0, style: { bold: true } },
        { sheetId: 's1', row: 1, column: 0, style: { bold: true, fontColor: '#0000FF' } },
      ]),
    )
    expect(store.bake('u1')).toHaveLength(2)
  })
})

describe('withBakedFormulaFormats', () => {
  it('merges into the user edit of the same cell and adds valid style-only edits', () => {
    const edits = withBakedFormulaFormats(
      [
        {
          sheetId: 's1',
          row: 0,
          column: 0,
          writeValue: true,
          value: 'x',
          style: { italic: true, bold: false },
        },
      ],
      [
        { sheetId: 's1', row: 0, column: 0, style: { bold: true } },
        { sheetId: 's1', row: 2, column: 3, style: { fillColor: '#FFFF00' } },
      ],
    )
    expect(edits[0]).toEqual({
      sheetId: 's1',
      row: 0,
      column: 0,
      writeValue: true,
      value: 'x',
      style: { italic: true, bold: true },
    })
    expect(edits[1]).toEqual({
      sheetId: 's1',
      row: 2,
      column: 3,
      writeValue: false,
      value: null,
      style: { fillColor: '#FFFF00' },
    })
    for (const edit of edits) expect(workbookCellEditSchema.safeParse(edit).success).toBe(true)
  })
})

describe('composeFormattedStyle', () => {
  it('overlays the patch on the cell style; false/null attributes are removed', () => {
    const composed = composeFormattedStyle(
      {
        bl: 1,
        it: 1,
        bd: { t: { s: 1, cl: { rgb: '#000000' } }, b: { s: 1, cl: { rgb: '#000000' } } },
      },
      { italic: false, fillColor: '#FF0000', borderBottom: null },
    )
    expect(composed).toMatchObject({ bl: 1, bg: { rgb: '#FF0000' } })
    expect(composed).not.toHaveProperty('it')
    expect(composed.bd).toHaveProperty('t')
    expect(composed.bd).not.toHaveProperty('b')
  })
})

describe('DVH format helpers', () => {
  it.each([
    ['B', { bold: true }, ['Đậm']],
    ['bi', { bold: true, italic: true }, ['Đậm', 'Nghiêng']],
    ['U', { underline: true, underlineStyle: 'single' }, ['Gạch chân đơn']],
    ['BU2', { underline: true, underlineStyle: 'double', bold: true }, ['Gạch chân kép', 'Đậm']],
    ['N', { bold: false, italic: false, underline: false }, ['Bình thường']],
  ])('DVH.Font codes %s set (not toggle) the styles', (codes, style, labels) => {
    expect(parseDvhFontCodes(codes)).toEqual({ style, labels })
  })

  it('rejects codes without a known style', () => {
    expect(parseDvhFontCodes('XYZ')).toBeNull()
    expect(parseDvhFontCodes('  ')).toBeNull()
  })

  it('converts Excel OLE colors (BGR) to #RRGGBB like the add-in', () => {
    expect(oleColorToHex(255)).toBe('#FF0000')
    expect(oleColorToHex(65535)).toBe('#FFFF00')
    expect(oleColorToHex(16711680)).toBe('#0000FF')
    expect(oleColorToHex(33023)).toBe('#FF8000')
    expect(oleColorToHex(-5)).toBe('#000000')
  })

  it('formats A1 addresses', () => {
    expect(a1Address({ startRow: 0, endRow: 9, startColumn: 0, endColumn: 2 })).toBe('A1:C10')
    expect(a1Address({ startRow: 4, endRow: 4, startColumn: 27, endColumn: 27 })).toBe('AB5')
  })

  it('treats blank, FALSE, 0 and "0" as rejecting a row, like IsFalse', () => {
    for (const v of [null, '', ' ', false, 0, '0', 'false']) expect(isFalseLike(v)).toBe(true)
    for (const v of [true, 1, -2, 'Đạt', '#tag']) expect(isFalseLike(v)).toBe(false)
  })

  it('reads title_rows as a count or a legacy title range', () => {
    expect(parseTitleRows(2)).toBe(2)
    expect(parseTitleRows(null)).toBe(0)
    expect(parseTitleRows([[3]])).toBe(3)
    expect(parseTitleRows([['Tiêu đề']])).toBe(1)
    expect(
      parseTitleRows([
        ['a', 'b'],
        ['c', 'd'],
      ]),
    ).toBe(2)
    expect(parseTitleRows(-1)).toMatch(/không âm/)
  })

  it('selects body rows from full-height, body-height and scalar filters', () => {
    // 1 title row + 4 body rows
    // full height: the title row's own value is skipped; body rows 0 and 2 pass
    expect(selectTableRows(5, 1, [[[false], [true], [false], [true], [false]]])).toEqual([0, 2])
    expect(selectTableRows(5, 1, [[[1, 0, 1, 1]], 'x'])).toEqual([0, 2, 3])
    expect(
      selectTableRows(5, 1, [
        [['h'], ['Đạt'], [''], ['Đạt'], ['Đạt']],
        [[true], [true], [true], [false], [true]],
      ]),
    ).toEqual([0, 3])
    expect(selectTableRows(5, 1, [false])).toEqual([])
    expect(selectTableRows(5, 1, [])).toEqual([0, 1, 2, 3])
    expect(selectTableRows(5, 1, [[[true], [true]]])).toMatch(/không khớp/)
    expect(
      selectTableRows(5, 1, [
        [
          [true, true],
          [true, true],
        ],
      ]),
    ).toMatch(/một hàng, một cột/)
  })

  it('copies a source style completely so destination formats never show through', () => {
    const style = copiedCellStyle({ bl: 1, bg: { rgb: '#D9E2F3' }, n: { pattern: '#,##0.000' } })
    expect(style).toMatchObject({
      bold: true,
      fillColor: '#D9E2F3',
      numberFormat: '#,##0.000',
      italic: false,
      underline: false,
    })
    expect(copiedCellStyle(null)).toMatchObject({
      bold: false,
      fillColor: null,
      numberFormat: 'General',
    })
  })
})

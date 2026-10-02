import { describe, expect, it } from 'vitest'
import { dvhShortName, formulaCallsDvh } from '../src/renderer/dvh-function-aliases'
import {
  cssColorToOle,
  explainFormula,
  explainValue,
  filterSheetNames,
  matchesCriterion,
  nameMatchesCategory,
  roundText,
  splitArguments,
} from '../src/renderer/dvh-workbook-functions'
import {
  DvhCommandQueue,
  absoluteAddress,
  highlightMatches,
  isValidDefinedName,
  joinRichParts,
  nameListPreview,
  namesToDelete,
  orderJoinCells,
  quoteSheetName,
  type CommandRequest,
} from '../src/renderer/dvh-formula-commands'

const plain = (text: string) => ({
  text,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
})

describe('DVH short names', () => {
  it('drops the DVH. prefix', () => {
    expect(dvhShortName('DVH.Table')).toBe('TABLE')
    expect(dvhShortName('DVH.Font.Color')).toBe('FONT.COLOR')
    expect(dvhShortName('SUM')).toBeNull()
    expect(dvhShortName('DVH.')).toBeNull()
  })
  it('recognises DVH calls anywhere in a formula', () => {
    expect(formulaCallsDvh('=IFERROR(DVH.ReadNumber(A1),"")')).toBe(true)
    expect(formulaCallsDvh('=SUM(A1:A3)')).toBe(false)
  })
})

describe('color functions', () => {
  it('converts CSS colors to OLE', () => {
    expect(cssColorToOle('#FF0000')).toBe(255)
    expect(cssColorToOle('#0000ff')).toBe(16_711_680)
    expect(cssColorToOle('#fff')).toBe(16_777_215)
    expect(cssColorToOle('rgb(0, 255, 0)')).toBe(65_280)
    expect(cssColorToOle(undefined)).toBeNull()
  })
})

describe('DVH.Explain', () => {
  const sheet: Record<string, (string | number | null)[]> = {
    A1: [1.234],
    A2: [2],
    A3: ['x'],
    'A1:A3': [1.234, 2, 'x'],
    B1: ['a'],
    'B1:B3': ['a', 'b', 'a'],
    'C1:C3': [10, 20, 30],
    'Sheet2!A1': [7],
  }
  const read = (ref: string) => sheet[ref] ?? null
  it('spells out SUM and PRODUCT ranges', () => {
    expect(explainFormula('=SUM(A1:A3)', 2, read)).toBe('=1.23+2')
    expect(explainFormula('=PRODUCT(A1,A2)', 1, read)).toBe('=1.2*2')
  })
  it('keeps only SUMIF cells meeting the criterion', () => {
    expect(explainFormula('=SUMIF(B1:B3,"a",C1:C3)', 2, read)).toBe('=10+30')
    expect(explainFormula('=SUMIF(B1:B3,"z",C1:C3)', 2, read)).toBe('=0')
  })
  it('swaps single-cell references outside quoted text', () => {
    expect(explainFormula('=$A$1*2+A2-Sheet2!A1&"A1"', 0, read)).toBe('=1*2+2-7&"A1"')
    expect(explainFormula('=ROUND(A1,1)+SUM(A1:A3)*2', 2, read)).toBe('=ROUND(1.23,1)+SUM(A1:A3)*2')
  })
  it('explains static values like the add-in', () => {
    expect(explainValue(null, 2)).toBe('(Trống)')
    expect(explainValue(3.14159, 2)).toBe('Giá trị: 3.14')
    expect(explainValue('abc', 2)).toBe('Nội dung: abc')
    expect(roundText(-0.0001, 2)).toBe('0')
  })
  it('splits arguments and evaluates criteria', () => {
    expect(splitArguments('A1, "x,y", SUM(B1,B2)')).toEqual(['A1', '"x,y"', 'SUM(B1,B2)'])
    expect(matchesCriterion(5, '>=5')).toBe(true)
    expect(matchesCriterion('A', '<>a')).toBe(false)
    expect(matchesCriterion('A', 'a')).toBe(true)
  })
})

describe('names and sheets', () => {
  const names = [
    { name: 'Gia', ref: 'Sheet1!$A$1', hidden: false },
    { name: 'Loi', ref: '#REF!', hidden: false },
    { name: 'Ngoai', ref: '[Book2.xlsx]Sheet1!$A$1', hidden: true },
    { name: 'gia_cu', ref: 'Sheet1!$B$1', hidden: false },
  ]
  it('counts by category', () => {
    expect(names.filter((n) => nameMatchesCategory(n, 'ALL'))).toHaveLength(4)
    expect(names.filter((n) => nameMatchesCategory(n, 'err')).map((n) => n.name)).toEqual(['Loi'])
    expect(names.filter((n) => nameMatchesCategory(n, 'OUT')).map((n) => n.name)).toEqual(['Ngoai'])
    expect(names.filter((n) => nameMatchesCategory(n, 'HIDDEN')).map((n) => n.name)).toEqual([
      'Ngoai',
    ])
    expect(names.filter((n) => nameMatchesCategory(n, 'XYZ'))).toHaveLength(0)
  })
  it('picks names to delete by category, name or wildcard', () => {
    expect(namesToDelete(names, 'gia*')).toEqual(['Gia', 'gia_cu'])
    expect(namesToDelete(names, 'GIA')).toEqual(['Gia'])
    expect(namesToDelete(names, 'ERR')).toEqual(['Loi'])
    expect(nameListPreview(['a', 'b', 'c', 'd', 'e', 'f', 'g'])).toBe('[a, b, c, d, e... +2 more]')
  })
  it('validates defined names as Excel does', () => {
    expect(isValidDefinedName('DonGia')).toBe(true)
    expect(isValidDefinedName('Đơn_giá.1')).toBe(true)
    expect(isValidDefinedName('X1')).toBe(false)
    expect(isValidDefinedName('R1C1')).toBe(false)
    expect(isValidDefinedName('1abc')).toBe(false)
    expect(isValidDefinedName('a b')).toBe(false)
  })
  it('quotes sheet names and builds absolute addresses', () => {
    expect(quoteSheetName('Sheet1')).toBe('Sheet1')
    expect(quoteSheetName("Kế hoạch's")).toBe("'Kế hoạch''s'")
    expect(absoluteAddress({ startRow: 0, endRow: 9, startColumn: 1, endColumn: 2 })).toBe(
      '$B$1:$C$10',
    )
  })
  it('lists sheets matching a filter', () => {
    expect(filterSheetNames(['KL1', 'Tong hop', 'kl2'], 'KL')).toEqual([
      "Sheets matching 'kl':",
      'KL1',
      'kl2',
    ])
    expect(filterSheetNames(['A'], '')).toEqual(['List of all sheets:', 'A'])
  })
})

describe('JoinFormat / FindFormat', () => {
  it('orders cells by the four join styles', () => {
    const cells = [
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 1, column: 0 },
    ]
    const at = (style: number) => orderJoinCells(cells, style).map((c) => `${c.row}${c.column}`)
    expect(at(1)).toEqual(['00', '01', '10'])
    expect(at(2)).toEqual(['10', '00', '01'])
    expect(at(3)).toEqual(['01', '00', '10'])
    expect(at(4)).toEqual(['10', '01', '00'])
  })
  it('joins rich parts with an unformatted delimiter, skipping empty cells', () => {
    const bold = { ...plain('A'), bold: true }
    const joined = joinRichParts(
      [
        { text: 'A', runs: [bold] },
        { text: '', runs: [] },
        { text: 'b', runs: [plain('b')] },
      ],
      ', ',
    )
    expect(joined.text).toBe('A, b')
    expect(joined.runs.map((r) => [r.text, r.bold])).toEqual([
      ['A', true],
      [', ', false],
      ['b', false],
    ])
  })
  it('formats every case-insensitive match on top of existing runs', () => {
    const result = highlightMatches(
      [{ ...plain('Bê tông '), italic: true }, plain('BÊ TÔNG')],
      ['bê'],
      {
        bold: true,
        color: '#FF0000',
      },
    )
    expect(result?.matches).toBe(2)
    expect(result?.runs.map((r) => [r.text, r.bold, r.italic, r.color ?? ''])).toEqual([
      ['Bê', true, true, '#FF0000'],
      [' tông ', false, true, ''],
      ['BÊ', true, false, '#FF0000'],
      [' TÔNG', false, false, ''],
    ])
    expect(highlightMatches([plain('abc')], ['z'], { bold: true })).toBeNull()
  })
})

describe('formula command queue', () => {
  const request = (key: string, row = 0): CommandRequest => ({
    unitId: 'u',
    sheetId: 's',
    row,
    column: 0,
    fn: 'DVH.ADDNAME',
    key,
    run: () => undefined,
  })
  it('only remembers file formulas seen in the first calculation', () => {
    const queue = new DvhCommandQueue()
    queue.request(request('a'))
    expect(queue.take(() => false)).toHaveLength(0)
    queue.markCalculated('u')
    queue.request(request('a'))
    expect(queue.take(() => false)).toHaveLength(0)
    queue.request(request('b'))
    expect(queue.take(() => false)).toHaveLength(1)
  })
  it('runs formulas the user typed, once per fingerprint', () => {
    const queue = new DvhCommandQueue()
    queue.request(request('a'))
    expect(queue.take(() => true)).toHaveLength(1)
    queue.request(request('a'))
    expect(queue.take(() => true)).toHaveLength(0)
    queue.forgetCell('u', 's', 0, 0)
    queue.request(request('a'))
    expect(queue.take(() => true)).toHaveLength(1)
    expect(queue.count('DVH.ADDNAME')).toBe(1)
    expect(queue.forget('DVH.ADDNAME')).toBe(1)
    expect(queue.count('DVH.ADDNAME')).toBe(0)
  })
})

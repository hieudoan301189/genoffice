import { describe, expect, it } from 'vitest'
import { DVH_PURE_FUNCTION_NAMES, evaluateDvhArray, evaluateDvhPure } from '../src/renderer/dvh-functions'

describe('DVH pure worksheet functions', () => {
  it('registers each supported function name once', () => {
    expect(DVH_PURE_FUNCTION_NAMES).toHaveLength(22)
    expect(new Set(DVH_PURE_FUNCTION_NAMES.map((name) => name.toUpperCase())).size).toBe(22)
  })

  it.each([
    ['DVH.ReadNumber', [123.32, 'Số tiền:', './.', 'đồng', 'hào xu'], 'Số tiền: Một trăm hai mươi ba đồng ba hào hai xu ./.'],
    ['DVH.ReadTime', ['14:30'], 'mười bốn giờ rưỡi'],
    ['DVH.UnMark', ['Tiếng Việt Đẹp'], 'Tieng Viet Dep'],
    ['DVH.FirstChar', ['Đại Việt Hùng'], 'ĐVH'],
    ['DVH.FirstChar', ['Đại Việt Hùng', false], 'ĐVH'],
    ['DVH.Roman', [2024], 'MMXXIV'],
    ['DVH.GetString', ['ABC123DEF'], 'ABCDEF'],
    ['DVH.GetNumber', ['Lô 12A, giá 500'], 12],
    ['DVH.IsInclude', ['Hello World', 'world'], true],
    ['DVH.IsInclude', ['Hello World', 'world', true], false],
    ['DVH.CountText', ['aaaa', 'aa'], 2],
    ['DVH.Evaluate', ['2*3*4*5'], '2*3*4*5=120'],
    ['DVH.Evaluate', ['-(2+3)^2'], '-(2+3)^2=-25'],
    ['DVH.GetText_XY', ['A[x]B[y]C', '[x]', '[y]'], 'B'],
    ['DVH.InsertText', ['ABCDE', '-', 3], 'AB-CDE'],
    ['DVH.NumToCol', [27], 'AA'],
    ['DVH.ColToNum', ['AA'], 27],
    ['DVH.RGBtoNum', [255, 128, 0], 33023],
    ['DVH.RGBtoHex', [255, 128, 0], '#FF8000'],
    ['DVH.NumToHex', [33023], '#FF8000'],
    ['DVH.WeekdayVN', [45292], 'Thứ Hai'],
    ['DVH.DayInMonth', [45292], 31],
    ['DVH.WeekInMonth', [45292], 5],
  ] as const)('%s(%j) = %j', (name, args, result) => {
    expect(evaluateDvhPure(name, args)).toBe(result)
  })

  it('retains DVH error and boundary behavior', () => {
    expect(evaluateDvhPure('DVH.Roman', [0])).toBe('#NUM! (Must be >= 1)')
    expect(evaluateDvhPure('DVH.NumToCol', [0])).toBe('#VALUE!')
    expect(evaluateDvhPure('DVH.RGBtoHex', [256, 0, 0])).toBe('#NUM!')
    expect(evaluateDvhPure('DVH.DayInMonth', [45351])).toBe(29)
    expect(evaluateDvhPure('DVH.DayInMonth', ['29/02/2024'])).toBe(29)
    expect(evaluateDvhPure('DVH.DayInMonth', ['30/02/2024'])).toBe('#VALUE!')
    expect(evaluateDvhPure('DVH.NumToCol', [1e100])).toBe('#VALUE!')
    expect(evaluateDvhPure('DVH.Evaluate', ['process.exit(1)'])).toBe('#VALUE!')
  })

  it('joins cells in DVH order and filters by include rows', () => {
    const values = [['A', 'B'], ['C', 'A']]
    expect(evaluateDvhArray('DVH.JoinText', values, [';', true, false, 0])).toBe('A;B;C')
    expect(evaluateDvhArray('DVH.JoinText', values, ['-', false, false, 3])).toBe('A-C-B-A')
    expect(evaluateDvhArray('DVH.JoinTextIF', [['A'], ['B'], ['C']], [';', false, false, false],
      [[true], [false], [true]])).toBe('C;A')
    expect(evaluateDvhArray('DVH.JoinTextIF', values, [], [[true]])).toContain('same row count')
  })
})

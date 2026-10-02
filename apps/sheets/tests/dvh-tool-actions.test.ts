/**
 * DVH Tool commands (DVH-Tool Ribbon/Etc handlers): text case with formula
 * wrapping, reference-style conversion outside text literals, adding and
 * removing ROUND, and the sheetView/@showZeros flag written on save.
 */
import { describe, expect, it } from 'vitest'
import {
  addRound,
  convertReferences,
  properCase,
  removeRound,
  transformCase,
  wrapCaseFormula,
} from '../src/renderer/dvh-tool-actions'
import { applyPageSetupState } from '@genoffice/xlsx-gateway/gateway/xlsx-page-setup'

describe('DVH Tool text case', () => {
  it('transforms values and wraps formulas like TextCaseHandler', () => {
    expect(transformCase('cầu rạch miễu', 'upper')).toBe('CẦU RẠCH MIỄU')
    expect(transformCase('CẦU Rạch', 'lower')).toBe('cầu rạch')
    expect(properCase('cầu RẠCH miễu-2b')).toBe('Cầu Rạch Miễu-2B')
    expect(transformCase('móng trụ T1', 'firstChar')).toBe('Móng trụ T1')
    expect(wrapCaseFormula('=A1&B1', 'upper')).toBe('=UPPER(A1&B1)')
    expect(wrapCaseFormula('=UPPER(A1)', 'upper')).toBeNull()
    expect(wrapCaseFormula('=A1', 'firstChar')).toBeNull()
  })
})

describe('DVH Tool reference style', () => {
  it('converts every A1 reference outside text, like ConvertFormula ToAbsolute', () => {
    const f = '=SUM(A1:B$2)+Sheet2!$C3&"A1"+VLOOKUP(d4,$E$1:F9,2,0)'
    expect(convertReferences(f, 'absolute')).toBe(
      '=SUM($A$1:$B$2)+Sheet2!$C$3&"A1"+VLOOKUP($D$4,$E$1:$F$9,2,0)',
    )
    expect(convertReferences(f, 'relative')).toBe(
      '=SUM(A1:B2)+Sheet2!C3&"A1"+VLOOKUP(D4,E1:F9,2,0)',
    )
    expect(convertReferences('=A1', 'absRow')).toBe('=A$1')
    expect(convertReferences('=A1', 'absCol')).toBe('=$A1')
    // function names with digits and long names are not references
    expect(convertReferences('=LOG10(A1)+ATAN2(B2,C3)', 'absolute')).toBe(
      '=LOG10($A$1)+ATAN2($B$2,$C$3)',
    )
    expect(convertReferences('text', 'absolute')).toBe('text')
  })
})

describe('DVH Tool ROUND', () => {
  it('adds and removes ROUND/ROUNDUP/ROUNDDOWN like RoundHandler', () => {
    expect(addRound({ formula: '=A1*B1' }, 'ROUND', 2)).toBe('=ROUND(A1*B1, 2)')
    expect(addRound({ value: 1.2345 }, 'ROUNDUP', 1)).toBe('=ROUNDUP(1.2345, 1)')
    expect(addRound({ value: 'abc' }, 'ROUND', 1)).toBeNull()
    expect(addRound({ formula: '=ROUND(A1,2)' }, 'ROUND', 2)).toBeNull()
    expect(removeRound('=ROUND(A1*B1, 2)')).toBe('=A1*B1')
    expect(removeRound('=ROUNDUP(SUM(A1,B1),0)+ROUNDDOWN(C1,1)')).toBe('=SUM(A1,B1)+C1')
    expect(removeRound('=ROUND(ROUND(A1,1)*2,2)')).toBe('=A1*2')
    expect(removeRound('=IF(A1>0,"ROUND(x,1)",0)')).toBe('=IF(A1>0,"ROUND(x,1)",0)')
  })
})

describe('hide zeros on save', () => {
  it('writes sheetView/@showZeros="0" and drops it again', () => {
    const xml =
      '<worksheet><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/></worksheet>'
    const hidden = applyPageSetupState(xml, { sheetName: 'S', showZeros: false })
    expect(hidden).toContain('showZeros="0"')
    expect(applyPageSetupState(hidden, { sheetName: 'S', showZeros: true })).not.toContain(
      'showZeros',
    )
  })
})

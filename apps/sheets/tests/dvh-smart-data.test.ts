import { describe, expect, it } from 'vitest'
import { emptyModel } from '@genoffice/dvh-model'
import {
  dvhPendingCount,
  inferFieldType,
  quoteSheetName,
  splitBindingFormula,
  type DvhSheetState,
} from '../src/renderer/dvh-smart-data'

describe('DVH Smart Data helpers (Sheets)', () => {
  it('splits binding formulas into sheet and cell', () => {
    expect(splitBindingFormula('Sheet1!$B$5')).toEqual({ sheet: 'Sheet1', cell: 'B5' })
    expect(splitBindingFormula("'Dữ liệu'!$AB$12")).toEqual({ sheet: 'Dữ liệu', cell: 'AB12' })
    expect(splitBindingFormula("'O''Brien'!C3")).toEqual({ sheet: "O'Brien", cell: 'C3' })
    expect(splitBindingFormula('Sheet1!#REF!')).toBeNull()
    expect(splitBindingFormula('Sheet1!$A$1:$B$2')).toBeNull()
  })

  it('quotes sheet names the way Excel formulas need', () => {
    expect(quoteSheetName('Sheet1')).toBe('Sheet1')
    expect(quoteSheetName('Dữ liệu')).toBe("'Dữ liệu'")
    expect(quoteSheetName("O'Brien")).toBe("'O''Brien'")
  })

  it('infers a field type from the cell value', () => {
    expect(inferFieldType(12.5)).toBe('number')
    expect(inferFieldType(true)).toBe('boolean')
    expect(inferFieldType('Dự án A')).toBe('text')
    expect(inferFieldType(null)).toBe('text')
  })

  it('counts a pending DVH change once', () => {
    const state: DvhSheetState = {
      model: emptyModel('doc_aaaaaaaaaaaaaaaa'),
      modelStoreItemId: null,
      history: null,
      historyStoreItemId: null,
      pending: [],
      warnedBroken: new Set(),
    }
    expect(dvhPendingCount(state)).toBe(0)
    state.pending.push({
      id: 'cs_1',
      txId: 'tx_1',
      docId: state.model!.docId,
      at: '2026-10-02T00:00:00.000Z',
      source: 'ui',
      action: 'Spreadsheet.BindField',
      changes: [],
    })
    expect(dvhPendingCount(state)).toBe(1)
    expect(dvhPendingCount(null)).toBe(0)
  })
})

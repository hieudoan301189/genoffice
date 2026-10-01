import { describe, expect, it } from 'vitest'
import { ObjectMatrix, type ICellData } from '@univerjs/core'
import type { BaseReferenceObject } from '@univerjs/engine-formula'
import { dvhLastUsed } from '../src/renderer/dvh-last-used'

function reference(cells: Record<number, Record<number, ICellData>>, row: number, column: number,
  changes: Record<number, Record<number, ICellData>> = {}): BaseReferenceObject {
  const model = new ObjectMatrix(cells)
  const runtime = new ObjectMatrix(changes)
  return {
    getRangePosition: () => ({ startRow: row, endRow: row, startColumn: column, endColumn: column }),
    getCurrentActiveSheetData: () => ({ cellData: model }),
    getCurrentRuntimeSheetData: () => runtime,
    getCurrentActiveArrayFormulaCellData: () => undefined,
    getCurrentRuntimeActiveArrayFormulaCellData: () => undefined,
    getCellData: (r: number, c: number) => runtime.getValue(r, c) ?? model.getValue(r, c),
  } as unknown as BaseReferenceObject
}

describe('DVH.LastRow and DVH.LastCol', () => {
  const cells: Record<number, Record<number, ICellData>> = {
    0: { 0: { v: 'A' }, 3: { v: 5 } },
    4: { 0: { s: 'style-2' }, 2: { f: '=1+2' } },
    11: { 0: { v: 'last' }, 1: { s: 'style-3' } },
  }

  it('finds the last occupied row in the first referenced column', () => {
    expect(dvhLastUsed(reference(cells, 2, 0), 'row')).toBe(12)
    expect(dvhLastUsed(reference(cells, 0, 2), 'row')).toBe(5)
    expect(dvhLastUsed(reference(cells, 0, 1), 'row')).toBe(1)
  })

  it('finds the last occupied column in the first referenced row', () => {
    expect(dvhLastUsed(reference(cells, 0, 0), 'column')).toBe(4)
    expect(dvhLastUsed(reference(cells, 4, 0), 'column')).toBe(3)
    expect(dvhLastUsed(reference(cells, 11, 0), 'column')).toBe(1)
  })

  it('uses live cells after edits instead of counting removed content', () => {
    const ref = reference(cells, 0, 0, { 11: { 0: { v: null } }, 8: { 0: { v: 'new' } } })
    expect(dvhLastUsed(ref, 'row')).toBe(9)
  })
})

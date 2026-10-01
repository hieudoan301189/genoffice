import { describe, expect, it } from 'vitest'
import {
  deleteDvhRows, insertDvhChild, lookupDvhMany, lookupDvhTwoKeys,
  sortDvhRows, subDvhArray, sumDvhExpression,
} from '../src/renderer/dvh-matrix-functions'

describe('DVH array functions from DVH-Excel', () => {
  const rows = [[3, 'C'], [1, 'A'], [2, 'B']]

  it('deletes and slices one-based rows without changing the source', () => {
    expect(deleteDvhRows(rows, 2, 2)).toEqual([[3, 'C'], [2, 'B']])
    expect(deleteDvhRows(rows, 1, 3)).toEqual([['#ERROR: Cannot delete all rows']])
    expect(subDvhArray(rows, 2, 1, 0, 2)).toEqual([[1, 'A'], [2, 'B']])
    expect(rows).toEqual([[3, 'C'], [1, 'A'], [2, 'B']])
  })

  it('inserts before, after, or over a row', () => {
    const child = [[9, 'X']]
    expect(insertDvhChild(rows, child, 2)).toEqual([[3, 'C'], [9, 'X'], [1, 'A'], [2, 'B']])
    expect(insertDvhChild(rows, child, 2, -1)).toEqual([[3, 'C'], [1, 'A'], [9, 'X'], [2, 'B']])
    expect(insertDvhChild(rows, child, 2, 0)).toEqual([[3, 'C'], [9, 'X'], [2, 'B']])
    expect(insertDvhChild(rows, [[9]], 2)).toEqual([['#ERROR: Column count mismatch (Parent: 2, Child: 1)']])
  })

  it('sorts numeric and text columns while retaining whole rows', () => {
    expect(sortDvhRows(rows, 1)).toEqual([[1, 'A'], [2, 'B'], [3, 'C']])
    expect(sortDvhRows(rows, 2, false, true)).toEqual([[3, 'C'], [2, 'B'], [1, 'A']])
    expect(sortDvhRows([['n/a'], [5], [2]], 1)).toEqual([['n/a'], [2], [5]])
  })

  it('returns multiple or two-condition matches', () => {
    const keys = [['A'], ['B'], ['A']]
    const results = [['one'], ['two'], ['one']]
    expect(lookupDvhMany('A', keys, results, ';')).toBe('one;one')
    expect(lookupDvhMany('A', keys, results, ';', true)).toBe('one')
    expect(lookupDvhTwoKeys(results, 'A', keys, 2, [[1], [2], [2]])).toBe('one')
    expect(lookupDvhTwoKeys(results, 'C', keys, 2, [[1], [2], [2]])).toBe('')
  })

  it('formats the DVH.Sum expression and skips nonnumeric cells', () => {
    expect(sumDvhExpression([[1, 'x'], [-2, '3'], [0, false]])).toBe('=1-2+3')
    expect(sumDvhExpression([['x', '']])).toBe('0')
  })
})

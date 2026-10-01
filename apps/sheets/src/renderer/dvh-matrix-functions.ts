// Worksheet-only ports of DVH-Excel/Function/clsFun_Array.cs.
// Inputs are snapshots of formula arguments; these helpers never edit a sheet.
export type DvhScalar = string | number | boolean
export type DvhMatrix = DvhScalar[][]

const error = (message: string): DvhMatrix => [[`#ERROR: ${message}`]]
const asText = (value: DvhScalar): string =>
  typeof value === 'boolean' ? (value ? 'True' : 'False') : String(value)

function asNumber(value: DvhScalar): number | null {
  if (typeof value === 'boolean' || value === '') return null
  const result = typeof value === 'number' ? value : Number(value.trim().replace(',', '.'))
  return Number.isFinite(result) ? result : null
}

export function deleteDvhRows(values: DvhMatrix, begin: number, end: number): DvhMatrix {
  const rows = values.length
  if (begin < 1 || begin > rows || !Number.isInteger(begin))
    return error(`IndexBegin must be 1-${rows}`)
  if (end < begin || end > rows || !Number.isInteger(end))
    return error(`IndexEnd must be ${begin}-${rows}`)
  if (end - begin + 1 === rows) return error('Cannot delete all rows')
  return values.filter((_, index) => index < begin - 1 || index >= end)
}

export function subDvhArray(values: DvhMatrix, startRow: number, startColumn: number,
  numRows: number, numColumns: number): DvhMatrix {
  const rows = values.length, columns = values[0]?.length ?? 0
  if (!Number.isInteger(startRow) || startRow < 1 || startRow > rows)
    return error(`startRow must be 1-${rows}`)
  if (!Number.isInteger(startColumn) || startColumn < 1 || startColumn > columns)
    return error(`startColumn must be 1-${columns}`)
  const rowCount = numRows <= 0 ? rows - startRow + 1 : Math.min(Math.trunc(numRows), rows - startRow + 1)
  const columnCount = numColumns <= 0 ? columns - startColumn + 1 : Math.min(Math.trunc(numColumns), columns - startColumn + 1)
  return values.slice(startRow - 1, startRow - 1 + rowCount)
    .map((row) => row.slice(startColumn - 1, startColumn - 1 + columnCount))
}

export function insertDvhChild(parent: DvhMatrix, child: DvhMatrix,
  position: number, mode = 1): DvhMatrix {
  const parentColumns = parent[0]?.length ?? 0, childColumns = child[0]?.length ?? 0
  if (parentColumns !== childColumns)
    return error(`Column count mismatch (Parent: ${parentColumns}, Child: ${childColumns})`)
  if (![1, 0, -1].includes(mode) || !Number.isInteger(mode))
    return error(`Invalid insertMode: ${mode}`)
  if (!Number.isInteger(position)) return error('Invalid insertPosition')
  const start = Math.max(0, Math.min(parent.length, position - 1 + (mode === -1 ? 1 : 0)))
  if (mode !== 0) return [...parent.slice(0, start), ...child, ...parent.slice(start)]
  const replaceAt = Math.min(start, parent.length - 1)
  return [...parent.slice(0, replaceAt), ...child, ...parent.slice(replaceAt + child.length)]
}

export function sortDvhRows(values: DvhMatrix, column: number,
  ascending = true, textMode = false, ignoreCase = true): DvhMatrix {
  const columns = values[0]?.length ?? 0
  if (!Number.isInteger(column) || column < 1 || column > columns)
    return error(`Column must be 1-${columns}, got ${column}`)
  const index = column - 1
  return values.map((row, original) => ({ row, original })).sort((a, b) => {
    let comparison: number
    if (textMode) {
      let left = asText(a.row[index] ?? ''), right = asText(b.row[index] ?? '')
      if (ignoreCase) { left = left.toUpperCase(); right = right.toUpperCase() }
      comparison = left < right ? -1 : left > right ? 1 : 0
    } else {
      const left = asNumber(a.row[index] ?? '') ?? Number.NEGATIVE_INFINITY
      const right = asNumber(b.row[index] ?? '') ?? Number.NEGATIVE_INFINITY
      comparison = left < right ? -1 : left > right ? 1 : 0
    }
    return (ascending ? comparison : -comparison) || a.original - b.original
  }).map((item) => item.row)
}

export function lookupDvhMany(lookup: DvhScalar, keys: DvhMatrix, results: DvhMatrix,
  delimiter = ', ', unique = false): string {
  if (results.length < keys.length) return '#N/A!'
  const found: string[] = [], seen = new Set<string>()
  for (let index = 0; index < keys.length; index++) {
    if (asText(keys[index]?.[0] ?? '') !== asText(lookup)) continue
    const value = asText(results[index]?.[0] ?? '')
    if (unique && seen.has(value)) continue
    seen.add(value)
    found.push(value)
  }
  return found.join(delimiter)
}

export function lookupDvhTwoKeys(results: DvhMatrix, key1: DvhScalar, array1: DvhMatrix,
  key2: DvhScalar, array2: DvhMatrix): DvhScalar {
  if (array1.length < results.length || array2.length < results.length) return '#N/A!'
  for (let index = 0; index < results.length; index++) {
    if (asText(array1[index]?.[0] ?? '') === asText(key1) &&
      asText(array2[index]?.[0] ?? '') === asText(key2))
      return results[index]?.[0] ?? ''
  }
  return ''
}

export function sumDvhExpression(values: DvhMatrix): string {
  const numbers = values.flat().map(asNumber).filter((value): value is number => value !== null && value !== 0)
  if (!numbers.length) return '0'
  return '=' + numbers.map((value, index) => `${index && value > 0 ? '+' : ''}${value}`).join('')
}

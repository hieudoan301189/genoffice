import type { ICellData, Nullable, ObjectMatrix } from '@univerjs/core'
import type { BaseReferenceObject } from '@univerjs/engine-formula'

type CellMatrix = ObjectMatrix<Nullable<ICellData>> | undefined

/** Match Excel's End(xlUp/xlToLeft): inspect the entire referenced column/row. */
export function dvhLastUsed(reference: BaseReferenceObject, axis: 'row' | 'column'): number {
  const sheet = reference.getCurrentActiveSheetData()
  if (!sheet?.cellData) return 1
  const range = reference.getRangePosition()
  const target = axis === 'row' ? range.startColumn : range.startRow
  const matrices: CellMatrix[] = [
    sheet.cellData,
    reference.getCurrentRuntimeSheetData(),
    reference.getCurrentActiveArrayFormulaCellData(),
    reference.getCurrentRuntimeActiveArrayFormulaCellData(),
  ]
  let last = 0
  for (const matrix of matrices) {
    if (!matrix) continue
    if (axis === 'row') {
      matrix.forEach((row, columns) => {
        if (row < last || !Object.hasOwn(columns, target)) return
        if (hasContent(reference.getCellData(row, target))) last = row
      })
    } else {
      const columns = matrix.getRow(target)
      if (!columns) continue
      for (const key of Object.keys(columns)) {
        const column = Number(key)
        if (column < last) continue
        if (hasContent(reference.getCellData(target, column))) last = column
      }
    }
  }
  return last + 1
}

function hasContent(cell: Nullable<ICellData>): boolean {
  return cell != null && (cell.f != null || cell.v != null || cell.p != null)
}

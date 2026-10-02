/**
 * Hide zero values (DVH Tool "Ẩn số 0", Excel's Show a zero in cells that
 * have zero value), persisted per sheet as sheetView/@showZeros="0".
 *
 * Display only: a CELL_CONTENT interceptor shows numeric zeros as empty;
 * the model, formulas and saved values keep the zero.
 */
import { CellValueType, InterceptorEffectEnum, type ICellData } from '@univerjs/core'
import { IRenderManagerService } from '@univerjs/engine-render'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'
import { SheetSkeletonManagerService } from '@univerjs/sheets-ui'

import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/// Hidden-zero sheets of blank/demo workbooks (no LazyWorkbookState).
const demoHideZeroSheets = new Set<string>()

export function hideZeroSheets(state: LazyWorkbookState | null): Set<string> {
  return state ? state.hideZeroSheets : demoHideZeroSheets
}

const isZero = (cell: ICellData | null | undefined): boolean =>
  !!cell &&
  (cell.v === 0 || (cell.t === CellValueType.NUMBER && Number(cell.v) === 0 && cell.v !== ''))

export function installZeroViewInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  return injector.get(SheetInterceptorService).intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // above NUMFMT (10): a hidden zero never reaches the number format
    priority: 12,
    effect: InterceptorEffectEnum.Value,
    handler: (cell, location, next) => {
      if (
        !cell ||
        !hideZeroSheets(lazyWorkbookRef.current).has(location.subUnitId) ||
        !isZero(cell as ICellData)
      )
        return next(cell)
      return next({ ...cell, v: '', t: CellValueType.STRING, p: null })
    },
  })
}

/// Repaints a workbook after the set of hidden-zero sheets changed (skeletons cache cell text).
export function repaintZeroView(runtime: UniverRuntime, unitId: string): void {
  const render = runtime.univer.__getInjector().get(IRenderManagerService).getRenderById(unitId)
  render?.with(SheetSkeletonManagerService).reCalculate()
  render?.mainComponent?.makeDirty()
}

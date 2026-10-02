import { WrapStrategy, type ICellData, type IStyleData } from '@univerjs/core'
import type { RibbonCommandContext } from './ribbon-actions'
import { beginUndoBatch } from './op-executor'
import { changeTextIndent, type IndentPreset, type BorderPreset } from './dvh-home-presets'
import { shrinkToFitFontSize, readSheetRangeMapped } from './univer-sync'
import { lazySheetMeta } from './univer-state'
import { handleDvhTool, type DvhToolPayload } from './dvh-tool-actions'

/** Commands of dvh-tool-actions.ts (text case, visibility, formulas, visible copy). */
const DVH_TOOL_ACTIONS = new Set([
  'case',
  'hidden',
  'zero-hide',
  'ref-style',
  'round-add',
  'round-remove',
  'copy-visible',
  'paste-visible',
  'read-number',
])

// DVH-Excel Ribbon/AlignmentHandler.cs: toggle alignment and restore the prior style.
// Commands use the normal sheet mutations so undo, dirty tracking and save apply.
export async function handleDvhHome(ctx: RibbonCommandContext, encoded: string): Promise<void> {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!runtime || !sheet || !range) return
  if (ctx.lazyWorkbookRef.current && !ctx.lazyWorkbookRef.current.flags.preloadComplete) {
    ctx.setMessage('Đang tải bảng tính. Hãy chờ tải xong rồi áp dụng định dạng DVH.')
    return
  }
  if (range.getHeight() * range.getWidth() > 10000) {
    ctx.setMessage('Hãy chọn vùng tối đa 10.000 ô để áp dụng định dạng DVH.')
    return
  }
  const { action, payload } = JSON.parse(encoded) as {
    action: string
    payload?: IndentPreset | BorderPreset
  }
  if (DVH_TOOL_ACTIONS.has(action)) {
    await handleDvhTool(ctx, action, payload as unknown as DvhToolPayload)
    return
  }
  const fileFormulas = new Set<string>()
  const state = ctx.lazyWorkbookRef.current
  if (action.startsWith('indent-') && state) {
    const meta = lazySheetMeta(state, sheet.getSheetId())
    if (!meta) return
    const { startRow, endRow, startColumn, endColumn } = range.getRange()
    const read = await readSheetRangeMapped(
      state,
      sheet.getSheetId(),
      { startRow, endRow, startColumn, endColumn },
      meta,
    )
    if (ctx.lazyWorkbookRef.current !== state) return
    for (const cell of read?.screen.cells ?? [])
      if (cell.formula) fileFormulas.add(`${cell.row}:${cell.column}`)
  }
  const source = range.getCellDatas()
  const row0 = range.getRow(),
    col0 = range.getColumn()
  const anchorCustom = source[0]?.[0]?.custom ?? {}
  const off =
    action === 'center'
      ? anchorCustom.dvhCenterAcross === true
      : action === 'shrink'
        ? anchorCustom.dvhShrink === true
        : range.getCellStyleData()?.ht === 4
  const measure = document.createElement('canvas').getContext('2d')!
  const values: ICellData[][] = []
  for (let r = 0; r < range.getHeight(); r++) {
    const row: ICellData[] = []
    for (let c = 0; c < range.getWidth(); c++) {
      const cell = source[r]?.[c] ?? {}
      const style = sheet.getRange(row0 + r, col0 + c).getCellStyleData() ?? {}
      const custom: Record<string, unknown> = { ...cell.custom }
      let patch: ICellData = {}
      if (action.startsWith('indent-')) {
        // Like TabHandler, formulas and empty cells are never rewritten.
        const key = `${row0 + r}:${col0 + c}`
        const edit = state?.editJournal.cells.get(sheet.getSheetId())?.get(key)
        const hasFormula = Boolean(
          cell.f || cell.si || edit?.formula || (!edit?.hasValue && fileFormulas.has(key)),
        )
        if (!hasFormula && cell.v != null && cell.v !== '') {
          const text = changeTextIndent(
            String(cell.v),
            payload as IndentPreset,
            action.slice(7) as 'add' | 'remove' | 'clear',
          )
          if (text !== String(cell.v)) patch = { v: text, t: 1, p: null }
        }
      } else if (action === 'border') {
        const edges = (payload as BorderPreset).edges
        const border = (i: number) =>
          edges[i]!.style === 0 ? null : { s: edges[i]!.style, cl: { rgb: edges[i]!.color } }
        patch = {
          s: {
            bd: {
              t: border(r === 0 ? 0 : 5),
              b: border(r === range.getHeight() - 1 ? 1 : 5),
              l: border(c === 0 ? 2 : 4),
              r: border(c === range.getWidth() - 1 ? 3 : 4),
              tl_br: border(6),
              bl_tr: border(7),
            },
          } as IStyleData,
        }
      } else if (action === 'center' || action === 'justify') {
        const key = action === 'center' ? 'dvhBeforeCenter' : 'dvhBeforeJustify'
        const previous = custom[key] as { ht?: number; tb?: number } | undefined
        if (!off) custom[key] = { ht: style.ht ?? 0, tb: style.tb ?? WrapStrategy.OVERFLOW }
        if (action === 'center') custom.dvhCenterAcross = !off
        else custom.dvhCenterAcross = false
        patch = {
          s: {
            ht: off ? (previous?.ht ?? 0) : action === 'center' ? 2 : 4,
            tb: off
              ? (previous?.tb ?? WrapStrategy.OVERFLOW)
              : action === 'center'
                ? WrapStrategy.OVERFLOW
                : style.tb,
          } as IStyleData,
          custom,
        }
      } else if (action === 'shrink') {
        const previous = custom.dvhBeforeShrink as { fs: number; tb: number } | undefined
        const original = off
          ? (previous?.fs ??
            (typeof custom.dvhOriginalSize === 'number' ? custom.dvhOriginalSize : style.fs) ??
            11)
          : (style.fs ?? 11)
        const width = sheet.getSheet().getColumnWidth(col0 + c) - 8
        measure.font = `${style.it ? 'italic ' : ''}${style.bl ? 'bold ' : ''}${(original * 96) / 72}px "${style.ff ?? 'Arial'}"`
        const size = shrinkToFitFontSize(
          String(cell.v ?? ''),
          original,
          width,
          (text) => measure.measureText(text).width,
        )
        if (!off) custom.dvhBeforeShrink = { fs: original, tb: style.tb ?? WrapStrategy.OVERFLOW }
        custom.shrinkToFit = !off
        custom.dvhShrink = !off
        custom.dvhOriginalSize = original
        patch = {
          s: {
            fs: off ? original : (size ?? original),
            tb: off ? (previous?.tb ?? WrapStrategy.OVERFLOW) : WrapStrategy.CLIP,
          },
          custom,
        }
      } else if (action === 'fit') {
        patch = { s: { tb: WrapStrategy.WRAP } }
      }
      row.push(patch)
    }
    values.push(row)
  }
  const batch = beginUndoBatch(runtime)
  try {
    if (action === 'center' && !off) range.breakApart()
    if (action.startsWith('indent-')) {
      // An empty patch in a bulk setValues still rewrites a formula cell's
      // cached value in streamed mode. Write only changed text cells.
      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r]!.length; c++) {
          const patch = values[r]![c]!
          if ('v' in patch) sheet.getRange(row0 + r, col0 + c).setValues([[patch]])
        }
      }
    } else {
      range.setValues(values)
    }
    if (action === 'fit') {
      // Measure against each merged span instead of treating its anchor as one column.
      for (let r = 0; r < range.getHeight(); r++) {
        let height = 20
        for (let c = 0; c < range.getWidth(); c++) {
          const cell = source[r]?.[c]
          if (!cell || cell.v == null) continue
          const style = sheet.getRange(row0 + r, col0 + c).getCellStyleData() ?? {}
          const merged = sheet.getSheet().getMergedCell(row0 + r, col0 + c)
          if (merged && (merged.startRow !== row0 + r || merged.startColumn !== col0 + c)) continue
          const size = style.fs ?? 11
          measure.font = `${style.bl ? 'bold ' : ''}${(size * 96) / 72}px "${style.ff ?? 'Arial'}"`
          let width = 0
          for (
            let col = merged?.startColumn ?? col0 + c;
            col <= (merged?.endColumn ?? col0 + c);
            col++
          )
            width += sheet.getSheet().getColumnWidth(col)
          const lines = String(cell.v)
            .split('\n')
            .reduce(
              (n, line) =>
                n +
                Math.max(1, Math.ceil(measure.measureText(line).width / Math.max(1, width - 8))),
              0,
            )
          height = Math.max(
            height,
            Math.ceil(
              (((lines * size * 96) / 72) * 1.25 + 6) /
                (merged ? merged.endRow - merged.startRow + 1 : 1),
            ),
          )
        }
        sheet.setRowHeight(row0 + r, Math.min(546, height))
      }
    }
    ctx.setMessage('Đã áp dụng định dạng DVH cho vùng chọn. Ctrl+Z để hoàn tác.')
  } finally {
    batch.settle()
  }
}

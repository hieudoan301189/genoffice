/**
 * DVH Tool commands of the Home group and the DVH tab, ported in behavior
 * from DVH-Tool/Ribbon/Etc (TextCaseHandler, VisibilityHandler,
 * ReferenceStyleHandler, RoundHandler, CopyPasteHandler) and the ReadNumber
 * form. Every command edits through the normal sheet mutations inside one
 * undo batch, so Ctrl+Z undoes it as a whole and the edit journal saves it.
 * Where the add-in toggled through a private state cache, a second click here
 * is undone with Ctrl+Z instead (same result, one undo stack).
 */
import type { ICellData, IStyleData } from '@univerjs/core'
import type { RibbonCommandContext } from './ribbon-actions'
import { beginUndoBatch } from './op-executor'
import { t } from './i18n/locale'
import { hideZeroSheets, repaintZeroView } from './zero-view'
import { isSheetRemoved, journalSize, recordPageSetup } from './edit-journal'

// ---------------- pure helpers ----------------

export type CaseKind = 'upper' | 'lower' | 'proper' | 'firstChar'

/** Excel PROPER: the first letter of every run of letters upper case, the rest lower case. */
export function properCase(text: string): string {
  let out = ''
  let previousLetter = false
  for (const ch of text.toLowerCase()) {
    const letter = /\p{L}/u.test(ch)
    out += letter && !previousLetter ? ch.toUpperCase() : ch
    previousLetter = letter
  }
  return out
}

export function transformCase(text: string, kind: CaseKind): string {
  switch (kind) {
    case 'upper':
      return text.toUpperCase()
    case 'lower':
      return text.toLowerCase()
    case 'proper':
      return properCase(text)
    case 'firstChar':
      return text ? text[0]!.toUpperCase() + text.slice(1) : text
  }
}

/** A formula wrapped in UPPER/LOWER/PROPER (=A1 → =UPPER(A1)); null when not applicable. */
export function wrapCaseFormula(formula: string, kind: CaseKind): string | null {
  const fn = { upper: 'UPPER', lower: 'LOWER', proper: 'PROPER', firstChar: '' }[kind]
  if (!fn || !formula.startsWith('=')) return null
  if (formula.toUpperCase().startsWith(`=${fn}(`)) return null
  return `=${fn}(${formula.slice(1)})`
}

/** Splits a formula into code and "string literal" segments. */
function segments(formula: string): { text: string; quoted: boolean }[] {
  const out: { text: string; quoted: boolean }[] = []
  let i = 0
  let start = 0
  while (i < formula.length) {
    if (formula[i] === '"') {
      if (i > start) out.push({ text: formula.slice(start, i), quoted: false })
      let j = i + 1
      while (j < formula.length) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') {
            j += 2
            continue
          }
          break
        }
        j++
      }
      out.push({ text: formula.slice(i, j + 1), quoted: true })
      i = j + 1
      start = i
      continue
    }
    i++
  }
  if (start < formula.length) out.push({ text: formula.slice(start), quoted: false })
  return out
}

export type RefKind = 'absolute' | 'absRow' | 'absCol' | 'relative'

const CELL_REF = /(?<![A-Za-z0-9_.$'])(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})(?![A-Za-z0-9_(])/g

/** Excel's ConvertFormula ToAbsolute: every A1 reference outside text takes the given kind. */
export function convertReferences(formula: string, kind: RefKind): string {
  if (!formula.startsWith('=')) return formula
  return segments(formula)
    .map((seg) =>
      seg.quoted
        ? seg.text
        : seg.text.replace(CELL_REF, (_m, _c, col: string, _r, row: string) => {
            const C = col.toUpperCase()
            if (Number(row) < 1) return `${_c}${col}${_r}${row}`
            switch (kind) {
              case 'absolute':
                return `$${C}$${row}`
              case 'absRow':
                return `${C}$${row}`
              case 'absCol':
                return `$${C}${row}`
              case 'relative':
                return `${C}${row}`
            }
          }),
    )
    .join('')
}

export type RoundFn = 'ROUND' | 'ROUNDUP' | 'ROUNDDOWN'

/** RoundHandler.AddRoundToCell: =f → =ROUND(f, n); a number → =ROUND(number, n). */
export function addRound(
  cell: { formula?: string | null; value?: unknown },
  fn: RoundFn,
  digits: number,
): string | null {
  if (cell.formula) {
    if (cell.formula.toUpperCase().includes(`${fn}(`)) return null
    const inner = cell.formula.startsWith('=') ? cell.formula.slice(1) : cell.formula
    return `=${fn}(${inner}, ${digits})`
  }
  const value = cell.value
  if (
    typeof value === 'number' ||
    (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value)))
  )
    return `=${fn}(${String(value).trim()}, ${digits})`
  return null
}

/** For each character of a formula: true inside a "text literal". */
function textMask(formula: string): boolean[] {
  const mask: boolean[] = []
  let pos = 0
  for (const seg of segments(formula)) {
    for (let i = 0; i < seg.text.length; i++) mask[pos + i] = seg.quoted
    pos += seg.text.length
  }
  return mask
}

/** RoundHandler.RemoveRoundFunctions: ROUND/ROUNDUP/ROUNDDOWN(x, n) → x, nested ones too. */
export function removeRound(formula: string): string {
  let result = formula
  for (let guard = 0; guard < 100; guard++) {
    // the first ROUND( outside a text literal ("ROUND(x)" in quotes stays text)
    const quoted = textMask(result)
    const re = /\b(ROUND|ROUNDUP|ROUNDDOWN)\s*\(/gi
    let m: RegExpExecArray | null
    while ((m = re.exec(result)) && quoted[m.index]) {
      /* skip matches inside quotes */
    }
    if (!m) break
    const open = m.index + m[0].length - 1
    let depth = 0
    let close = -1
    let lastComma = -1
    let inText = false
    for (let i = open; i < result.length; i++) {
      const ch = result[i]
      if (ch === '"') inText = !inText
      if (inText) continue
      if (ch === '(') depth++
      else if (ch === ')') {
        depth--
        if (depth === 0) {
          close = i
          break
        }
      } else if (ch === ',' && depth === 1) lastComma = i
    }
    if (close < 0 || lastComma < 0) break
    const inner = result.slice(open + 1, lastComma).trim()
    result = result.slice(0, m.index) + inner + result.slice(close + 1)
  }
  return result
}

// ---------------- commands ----------------

/** What the copy-visible command keeps for paste-visible: rows of cells (values, formulas, styles). */
let visibleClipboard: ICellData[][] | null = null

export type DvhToolPayload =
  | { kind: CaseKind }
  | { ref: RefKind }
  | { fn: RoundFn; digits: number }
  | {
      input: string
      output: string
      before: string
      after: string
      intUnit: string
      decUnit: string
    }
  | undefined

const A1 = (row: number, col: number) => {
  let n = col + 1
  let letters = ''
  while (n > 0) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters
    n = Math.floor((n - 1) / 26)
  }
  return `${letters}${row + 1}`
}
const quote = (text: string) => `"${text.replace(/"/g, '""')}"`

export async function handleDvhTool(
  ctx: RibbonCommandContext,
  action: string,
  payload: DvhToolPayload,
): Promise<void> {
  const runtime = ctx.univerRef.current
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!runtime || !workbook || !sheet || !range) return

  if (action === 'zero-hide') {
    const state = ctx.lazyWorkbookRef.current
    const sheetId = sheet.getSheetId()
    const set = hideZeroSheets(state)
    const hide = !set.has(sheetId)
    if (hide) set.add(sheetId)
    else set.delete(sheetId)
    repaintZeroView(runtime, workbook.getId())
    if (state && !isSheetRemoved(state.editJournal, sheetId)) {
      recordPageSetup(state.editJournal, sheetId, { showZeros: !hide })
      ctx.setPendingEdits(journalSize(state.editJournal))
    }
    ctx.setMessage(t(hide ? 'dvhToolZerosHidden' : 'dvhToolZerosShown'))
    return
  }

  if (action === 'paste-visible') {
    if (!visibleClipboard) {
      ctx.setMessage(t('dvhToolNothingCopied'))
      return
    }
    // PasteVisible: each source row lands on the next visible row from the target
    const ws = sheet.getSheet()
    let row = range.getRow()
    const col0 = range.getColumn()
    const batch = beginUndoBatch(runtime)
    try {
      for (const cells of visibleClipboard) {
        while (row < ws.getRowCount() && !ws.getRowVisible(row)) row++
        let col = col0
        for (const cell of cells) {
          while (col < ws.getColumnCount() && !ws.getColVisible(col)) col++
          sheet.getRange(row, col).setValues([[cell]])
          col++
        }
        row++
      }
    } finally {
      batch.settle()
    }
    ctx.setMessage(t('dvhToolPasted', { count: visibleClipboard.length }))
    return
  }

  if (action === 'read-number') {
    const p = payload as Extract<DvhToolPayload, { input: string }>
    // empty fields mean the selected cell and the one to its right
    const inputRef = p.input.trim() || A1(range.getRow(), range.getColumn())
    const outputRef = p.output.trim() || A1(range.getRow(), range.getColumn() + 1)
    const target = sheet.getRange(outputRef)
    const source = inputRef
    // a cell reference stays live; a typed number is read as is
    const number = /^\$?[A-Za-z]{1,3}\$?\d+$/.test(source) ? source : quote(source)
    const formula = `=DVH.ReadNumber(${number}, ${quote(p.before)}, ${quote(p.after)}, ${quote(p.intUnit)}, ${quote(p.decUnit)})`
    const batch = beginUndoBatch(runtime)
    try {
      target.setValues([[{ f: formula }]])
    } finally {
      batch.settle()
    }
    ctx.setMessage(t('dvhToolReadNumberDone', { cell: outputRef.toUpperCase() }))
    return
  }

  const cells = range.getCellDatas()
  const row0 = range.getRow()
  const col0 = range.getColumn()
  const ws = sheet.getSheet()

  if (action === 'copy-visible') {
    const rows: ICellData[][] = []
    for (let r = 0; r < range.getHeight(); r++) {
      if (!ws.getRowVisible(row0 + r)) continue
      const row: ICellData[] = []
      for (let c = 0; c < range.getWidth(); c++) {
        if (!ws.getColVisible(col0 + c)) continue
        const cell = cells[r]?.[c] ?? {}
        const style = sheet.getRange(row0 + r, col0 + c).getCellStyleData()
        row.push({
          ...(cell.f ? { f: cell.f } : { v: cell.v ?? null }),
          ...(style ? { s: style as IStyleData } : {}),
        })
      }
      rows.push(row)
    }
    visibleClipboard = rows
    // plain text too, for pasting into other programs
    const tsv = rows
      .map((row) => row.map((cell) => String(cell.f ?? cell.v ?? '')).join('\t'))
      .join('\n')
    try {
      await navigator.clipboard.writeText(tsv)
    } catch {
      // the internal copy is enough for Paste visible
    }
    ctx.setMessage(t('dvhToolCopied', { count: rows.length }))
    return
  }

  let changed = 0
  const writes: { row: number; col: number; cell: ICellData }[] = []
  for (let r = 0; r < range.getHeight(); r++) {
    for (let c = 0; c < range.getWidth(); c++) {
      const cell = cells[r]?.[c] ?? {}
      const at = { row: row0 + r, col: col0 + c }
      const formula = typeof cell.f === 'string' && cell.f ? cell.f : null
      if (action === 'case') {
        const kind = (payload as { kind: CaseKind }).kind
        if (formula) {
          const wrapped = wrapCaseFormula(formula, kind)
          if (wrapped) writes.push({ ...at, cell: { f: wrapped } })
        } else if (typeof cell.v === 'string' && cell.v) {
          const text = transformCase(cell.v, kind)
          if (text !== cell.v) writes.push({ ...at, cell: { v: text, p: null } })
        }
      } else if (action === 'hidden') {
        const style = sheet.getRange(at.row, at.col).getCellStyleData()
        const pattern = style?.n?.pattern ?? ''
        const custom = { ...(cell.custom ?? {}) } as Record<string, unknown>
        const hidden = pattern === ';;;'
        const restore = typeof custom.dvhBeforeHidden === 'string' ? custom.dvhBeforeHidden : ''
        if (!hidden) custom.dvhBeforeHidden = pattern
        else delete custom.dvhBeforeHidden
        writes.push({
          ...at,
          cell: { s: { n: { pattern: hidden ? restore : ';;;' } } as IStyleData, custom },
        })
      } else if (action === 'ref-style' && formula) {
        const next = convertReferences(formula, (payload as { ref: RefKind }).ref)
        if (next !== formula) writes.push({ ...at, cell: { f: next } })
      } else if (action === 'round-add') {
        const { fn, digits } = payload as { fn: RoundFn; digits: number }
        const next = addRound({ formula, value: cell.v }, fn, digits)
        if (next) writes.push({ ...at, cell: { f: next } })
      } else if (action === 'round-remove' && formula) {
        const next = removeRound(formula)
        if (next !== formula) writes.push({ ...at, cell: { f: next } })
      }
    }
  }
  if (writes.length === 0) {
    ctx.setMessage(
      t(
        action === 'ref-style' || action === 'round-remove'
          ? 'dvhToolNoFormula'
          : 'dvhToolNothingChanged',
      ),
    )
    return
  }
  const batch = beginUndoBatch(runtime)
  try {
    for (const w of writes) {
      sheet.getRange(w.row, w.col).setValues([[w.cell]])
      changed++
    }
  } finally {
    batch.settle()
  }
  ctx.setMessage(t('dvhToolChanged', { count: changed }))
}

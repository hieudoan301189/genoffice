/**
 * DVH.Table in Docs (P2). A table renders a collection copied into the
 * document model through a workbook link: a regular docx table (header rows
 * repeat on every page, page-wide widths) inside a block content control
 * tagged `dvh:t:<tableId>`, so Refresh can find and rebuild it and Word shows
 * an ordinary table.
 */
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  dvhTableId,
  type CellBorders,
  type TableCell,
  type TableModel,
} from '@genoffice/docx-engine'
import {
  contentHash,
  formatCellValue,
  modelHash,
  newDvhId,
  renderTable,
  tableSdtPrXml,
  type DvhCellStyle,
  type DvhTable,
  type RenderedTable,
} from '@genoffice/dvh-model'
import { tableModelToPmNode } from './editor/convert'
import {
  ensureModel,
  recordDvhDocsChange,
  updateFromSource,
  type FieldConflict,
  type DvhDocsState,
  type DvhSource,
} from './dvh-smart-data'
import type { DvhLink } from '@genoffice/dvh-model'

const hex = (color: string | undefined) => color?.replace(/^#/, '').toUpperCase()

/** OS regional locale for number formats (set at startup); the UI language until it arrives */
let numberLocale: string | null = null
export function setDvhNumberLocale(locale: string | null): void {
  numberLocale = locale
}

function cellBorders(style: DvhCellStyle): CellBorders | undefined {
  if (!style.border) return undefined
  const side = {
    style: 'single',
    szEighths: style.border.width === 'medium' ? 12 : 4,
    color: hex(style.border.color)!,
  }
  return { top: side, left: side, bottom: side, right: side }
}

/** The docx table model of a rendered DVH table. */
export function renderedToTableModel(
  grid: RenderedTable,
  table: DvhTable,
  locale: string,
): TableModel {
  const rows: TableCell[][] = grid.rows.map((row) =>
    row.map((cell) => {
      const text = formatCellValue(cell.value, cell.numFmt, locale)
      const run = {
        text,
        ...(cell.style.bold ? { bold: true } : {}),
        ...(cell.style.italic ? { italic: true } : {}),
        ...(cell.style.color ? { color: hex(cell.style.color)! } : {}),
        ...(cell.style.fontSize ? { sizeHalfPoints: Math.round(cell.style.fontSize * 2) } : {}),
      }
      const borders = cellBorders(cell.style)
      return {
        paras: [text],
        richParas: [
          {
            runs: text ? [run] : [],
            ...(cell.style.align ? { align: cell.style.align } : {}),
          },
        ],
        ...(cell.colSpan > 1 ? { colSpan: cell.colSpan } : {}),
        ...(cell.style.fill ? { fill: hex(cell.style.fill)! } : {}),
        ...(cell.style.align ? { align: cell.style.align } : {}),
        ...(borders ? { borders } : {}),
      }
    }),
  )
  const none = { style: 'none' }
  return {
    rows,
    colWidthsPct: grid.widths.map((w) => w * 100),
    ...(table.layout.widths === 'page' ? { widthPct: 100, autoFit: 'window' as const } : {}),
    // cells carry their own borders; no table-level grid lines on top
    borders: { top: none, left: none, bottom: none, right: none, insideH: none, insideV: none },
    repeatHeaderRows: grid.rows.map((_, r) => table.layout.repeatHeader && r < grid.headerRowCount),
    rowCantSplit: grid.rows.map(() => table.layout.keepRowsTogether),
  }
}

/** Every rendered DVH table in the document. */
export function tableOccurrences(
  doc: PmNode,
): { tableId: string; sdtPr: string; pos: number; node: PmNode }[] {
  const out: { tableId: string; sdtPr: string; pos: number; node: PmNode }[] = []
  doc.forEach((node, pos) => {
    const sdtPr = node.type.name === 'docTable' ? (node.attrs.dvhTable as string | null) : null
    const tableId = sdtPr ? dvhTableId(sdtPr) : null
    if (sdtPr && tableId) out.push({ tableId, sdtPr, pos, node })
  })
  return out
}

/** What a table node shows, cell by cell: the hand-edit check compares it with lastRender. */
function nodeTextHash(node: PmNode): string {
  const rows: string[][] = []
  node.forEach((row) => {
    const cells: string[] = []
    row.forEach((cell) => cells.push(cell.textContent))
    rows.push(cells)
  })
  return contentHash(rows)
}

function usedSdtIds(doc: PmNode): Set<number> {
  const ids = new Set<number>()
  doc.descendants((node) => {
    for (const value of [node.attrs?.dvhTable, ...node.marks.map((m) => m.attrs.sdtPr)]) {
      const m = typeof value === 'string' ? /<w:id w:val="(-?\d+)"/.exec(value) : null
      if (m) ids.add(Number(m[1]))
    }
  })
  return ids
}

function tableNode(editor: Editor, dvh: DvhDocsState, table: DvhTable, sdtPr: string): PmNode {
  const grid = renderTable(table, dvh.model?.collections ?? [])
  const model = renderedToTableModel(grid, table, numberLocale || navigator.language || 'en-US')
  const json = tableModelToPmNode(model)
  return editor.schema.nodeFromJSON({ ...json, attrs: { ...json.attrs, dvhTable: sdtPr } })
}

/** A table definition showing every column of a collection in the document model. */
export function createDocTable(
  dvh: DvhDocsState,
  options: { collectionId: string; name?: string; mode?: 'source' | 'destination' | undefined },
  source: 'ui' | 'ai' = 'ui',
): DvhTable {
  const model = ensureModel(dvh).model
  const collection = model.collections.find(
    (c) => c.id === options.collectionId || c.name === options.collectionId,
  )
  if (!collection) throw new Error(`unknown collection ${options.collectionId}`)
  const taken = new Set(model.tables.map((tb) => tb.name))
  let name = options.name?.trim() || collection.name
  for (let n = 2; taken.has(name); n++) name = `${options.name?.trim() || collection.name} ${n}`
  const table: DvhTable = {
    id: newDvhId('t'),
    name,
    source: { collectionId: collection.id },
    columns: collection.columns.map((c) => ({
      id: newDvhId('c').replace(/^c_/, 'tc_'),
      columnId: c.id,
      title: c.title,
    })),
    autoColumns: true,
    style: { mode: options.mode ?? 'source' },
    layout: { repeatHeader: true, keepRowsTogether: true, widths: 'page' },
  }
  model.tables.push(table)
  recordDvhDocsChange(
    dvh,
    'Table.Create',
    [{ objectId: table.id, path: '', before: null, after: { name, collectionId: collection.id } }],
    source,
  )
  return table
}

/** Inserts a rendered table after the block holding the caret (or at `afterBlockIndex`). */
export function insertDocTable(
  editor: Editor,
  dvh: DvhDocsState,
  table: DvhTable,
  options: { afterBlockIndex?: number | undefined; source?: 'ui' | 'ai' } = {},
): void {
  let sdtId = 0
  const used = usedSdtIds(editor.state.doc)
  while (sdtId === 0 || used.has(sdtId)) sdtId = 1 + Math.floor(Math.random() * 2_000_000_000)
  const sdtPr = tableSdtPrXml({ tableId: table.id, alias: table.name, sdtId })
  const node = tableNode(editor, dvh, table, sdtPr)
  const { doc, selection } = editor.state
  let pos: number
  if (options.afterBlockIndex !== undefined) {
    if (options.afterBlockIndex < 0) pos = 0
    else {
      const index = Math.min(options.afterBlockIndex, doc.childCount - 1)
      pos = 0
      for (let i = 0; i <= index; i++) pos += doc.child(i).nodeSize
    }
  } else {
    pos = selection.$from.depth >= 1 ? selection.$from.after(1) : doc.content.size
  }
  editor.view.dispatch(editor.state.tr.insert(pos, node).scrollIntoView())
  table.lastRender = { hash: nodeTextHash(node), at: new Date().toISOString() }
  recordDvhDocsChange(
    dvh,
    'Document.InsertTable',
    [{ objectId: table.id, path: 'occurrence', before: null, after: table.name }],
    options.source ?? 'ui',
  )
}

export interface DocTableRefresh {
  /** occurrences rebuilt */
  readonly refreshed: number
  /** the table holds hand edits and `force` was not set: nothing changed */
  readonly needsConfirm: boolean
}

/** Rebuilds every occurrence of a table from the document model's collection. */
export function refreshDocTable(
  editor: Editor,
  dvh: DvhDocsState,
  tableRef: string,
  options: {
    force?: boolean | undefined
    source?: 'ui' | 'ai' | 'link' | 'restore'
    history?: boolean
  } = {},
): DocTableRefresh {
  const table = dvh.model?.tables.find((tb) => tb.id === tableRef || tb.name === tableRef)
  if (!table) throw new Error(`unknown table ${tableRef}`)
  const occurrences = tableOccurrences(editor.state.doc).filter((o) => o.tableId === table.id)
  if (
    !options.force &&
    table.lastRender &&
    occurrences.some((o) => nodeTextHash(o.node) !== table.lastRender!.hash)
  ) {
    return { refreshed: 0, needsConfirm: true }
  }
  const tr = editor.state.tr
  let hash: string | null = null
  // back to front, so earlier positions stay valid
  for (const occ of [...occurrences].reverse()) {
    const node = tableNode(editor, dvh, table, occ.sdtPr)
    hash = nodeTextHash(node)
    tr.replaceWith(occ.pos, occ.pos + occ.node.nodeSize, node)
  }
  if (options.history === false) tr.setMeta('addToHistory', false)
  if (tr.docChanged) editor.view.dispatch(tr)
  if (hash) table.lastRender = { hash, at: new Date().toISOString() }
  if (occurrences.length > 0) {
    recordDvhDocsChange(
      dvh,
      'Table.Refresh',
      [{ objectId: table.id, path: 'render', before: null, after: occurrences.length }],
      options.source ?? 'ui',
    )
  }
  return { refreshed: occurrences.length, needsConfirm: false }
}

/** Changes a table's style (render again to show it). */
export function setDocTableStyle(
  dvh: DvhDocsState,
  tableRef: string,
  style: { [K in keyof DvhTable['style']]?: DvhTable['style'][K] | undefined },
  source: 'ui' | 'ai' = 'ui',
): DvhTable {
  const table = dvh.model?.tables.find((tb) => tb.id === tableRef || tb.name === tableRef)
  if (!table) throw new Error(`unknown table ${tableRef}`)
  const before = table.style
  const defined = Object.fromEntries(Object.entries(style).filter(([, v]) => v !== undefined))
  table.style = { ...table.style, ...defined }
  recordDvhDocsChange(
    dvh,
    'Table.SetStyle',
    [{ objectId: table.id, path: 'style', before, after: table.style }],
    source,
  )
  return table
}

export interface LinkUpdate {
  readonly fields: number
  readonly tables: number
  /** tables edited by hand: not refreshed, the caller asks before forcing them */
  readonly editedTables: readonly string[]
  /** fields edited here only, waiting to be written back (none when the source did not change) */
  readonly localEdits?: readonly string[]
  /** fields edited on both sides (see FieldConflict) */
  readonly conflicts?: readonly FieldConflict[]
}

/**
 * "Update from source": fields, collections, then every table showing an
 * updated collection. `auto` (automatic links) keeps the changes out of the
 * undo history and does nothing when the source did not change.
 */
export function updateLinkFromSource(
  editor: Editor,
  dvh: DvhDocsState,
  link: DvhLink,
  source: DvhSource,
  options: { auto?: boolean } = {},
): LinkUpdate {
  if (options.auto && link.lastSync?.hash === modelHash(source.model)) {
    return { fields: 0, tables: 0, editedTables: [] }
  }
  const history = !options.auto
  const update = updateFromSource(editor, dvh, link, source, { history })
  let tables = 0
  const editedTables: string[] = []
  for (const table of dvh.model?.tables ?? []) {
    if (!('collectionId' in table.source)) continue
    if (!update.collections.includes(table.source.collectionId)) continue
    const result = refreshDocTable(editor, dvh, table.id, { source: 'link', history })
    if (result.needsConfirm) editedTables.push(table.id)
    else tables += result.refreshed
  }
  return {
    fields: update.fields,
    tables,
    editedTables,
    localEdits: update.localEdits,
    conflicts: update.conflicts,
  }
}

/**
 * Ends a link: its fields and tables stay in the document as plain content
 * (no content control), and the objects it brought leave the model.
 */
export function unlinkSource(editor: Editor, dvh: DvhDocsState, linkId: string): void {
  const model = dvh.model
  const link = model?.links.find((l) => l.id === linkId)
  if (!model || !link) return
  const targets = new Set(link.targets)
  const tables = model.tables.filter(
    (tb) => 'collectionId' in tb.source && targets.has(tb.source.collectionId),
  )
  const tableIds = new Set(tables.map((tb) => tb.id))
  const tr = editor.state.tr
  const markType = editor.schema.marks.dvhField!
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTable') {
      const id = node.attrs.dvhTable ? dvhTableId(String(node.attrs.dvhTable)) : null
      if (id && tableIds.has(id))
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, dvhTable: null })
      return true
    }
    if (!node.isText) return true
    for (const mark of node.marks) {
      if (mark.type !== markType) continue
      const fieldId = /<w:tag\s+w:val="dvh:f:([^"]+)"/.exec(String(mark.attrs.sdtPr ?? ''))?.[1]
      if (fieldId && targets.has(fieldId)) tr.removeMark(pos, pos + node.nodeSize, mark)
    }
    return true
  })
  if (tr.docChanged) editor.view.dispatch(tr)
  model.fields = model.fields.filter((f) => !targets.has(f.id))
  model.collections = model.collections.filter((c) => !targets.has(c.id))
  model.tables = model.tables.filter((tb) => !tableIds.has(tb.id))
  model.links = model.links.filter((l) => l.id !== linkId)
  recordDvhDocsChange(dvh, 'Link.Remove', [
    { objectId: link.id, path: '', before: link, after: null },
  ])
}

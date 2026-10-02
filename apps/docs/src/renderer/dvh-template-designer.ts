/**
 * Template Designer in Docs (P6): turns the open document into a Smart
 * Template. Sections are block-level content controls (kept by the docx
 * engine as block shells); column values are run-level controls (kept like
 * Smart Fields). Conditions live in the model, so the expression can be long.
 *
 * - `wrapSelection`: the top-level blocks of the selection become one section
 *   (`dvh:if:<id>`, `dvh:repeat:<collectionId>`, or `dvh:repeatrows:<collectionId>`
 *   when the selection is one table);
 * - `insertColumnControl`: a column value at the caret (`dvh:c:<columnId>`).
 */
import type { Editor } from '@tiptap/core'
import {
  diffModels,
  escapeXml,
  fieldText,
  newDvhId,
  type DvhCondition,
  type DvhModel,
} from '@genoffice/dvh-model'
import { exprError } from '@genoffice/dvh-template'
import {
  applyFieldText,
  ensureModel,
  recordDvhDocsChange,
  type DvhDocsState,
} from './dvh-smart-data'
import { refreshDocTable } from './dvh-tables'

export type SectionKind = 'if' | 'repeat'

const shellOpen = (tag: string, alias: string) =>
  `<w:sdt><w:sdtPr><w:alias w:val="${escapeXml(alias)}"/><w:tag w:val="${escapeXml(tag)}"/></w:sdtPr><w:sdtContent>`
const SHELL_CLOSE = '</w:sdtContent></w:sdt>'

/** The top-level block range of the selection: [first, last] child indexes. */
function selectedBlocks(editor: Editor): { first: number; last: number } {
  const { $from, $to } = editor.state.selection
  return { first: $from.index(0), last: Math.max($from.index(0), $to.index(0)) }
}

/** The section tag of the top-level block at `index`, or null. */
export function sectionTagAt(editor: Editor, index: number): string | null {
  const raw = editor.state.doc.maybeChild(index)?.attrs.sdtShell
  if (!raw) return null
  try {
    return (JSON.parse(String(raw)) as { tag?: string }).tag ?? null
  } catch {
    return null
  }
}

/**
 * Wraps the selected top-level blocks in a section control. Refused (false)
 * when a selected block already belongs to a content control: sections do
 * not nest at block level.
 */
export function wrapSelection(editor: Editor, tag: string, alias: string): boolean {
  const { first, last } = selectedBlocks(editor)
  const doc = editor.state.doc
  for (let i = first; i <= last; i++) if (doc.child(i).attrs.sdtShell) return false
  const tr = editor.state.tr
  let pos = 0
  for (let i = 0; i < first; i++) pos += doc.child(i).nodeSize
  for (let i = first; i <= last; i++) {
    const node = doc.child(i)
    if (!('sdtShell' in node.type.spec.attrs!)) return false
    const shell = {
      alias,
      tag,
      controlType: 'text',
      openXml: i === first ? shellOpen(tag, alias) : '',
      closeXml: i === last ? SHELL_CLOSE : '',
    }
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, sdtShell: JSON.stringify(shell) })
    pos += node.nodeSize
  }
  editor.view.dispatch(tr)
  return true
}

/** Adds (or updates) a named condition; returns it, or the syntax error. */
export function saveCondition(
  dvh: DvhDocsState,
  expr: string,
  name: string,
  id?: string,
): DvhCondition | { error: string } {
  const error = exprError(expr)
  if (error) return { error }
  const { model } = ensureModel(dvh)
  const conditions = (model.conditions ??= [])
  const existing = id ? conditions.find((c) => c.id === id) : undefined
  const before = existing ? { ...existing } : null
  const condition: DvhCondition = existing
    ? Object.assign(existing, { expr, ...(name ? { name } : {}) })
    : { id: newDvhId('r').replace(/^r_/, 'cond_'), expr, ...(name ? { name } : {}) }
  if (!existing) conditions.push(condition)
  recordDvhDocsChange(dvh, 'Template.SetCondition', [
    { objectId: condition.id, path: 'expr', before, after: { ...condition } },
  ])
  return condition
}

/** Wraps the selection in a condition section. */
export function wrapInCondition(editor: Editor, condition: DvhCondition): boolean {
  return wrapSelection(editor, `dvh:if:${condition.id}`, condition.name ?? condition.expr)
}

/**
 * Wraps the selection in a repeating section over a collection: a table
 * alone repeats its rows that hold column controls; anything else repeats as
 * a block. An optional condition filters the records.
 */
export function wrapInRepeat(
  editor: Editor,
  dvh: DvhDocsState,
  collectionId: string,
  conditionId?: string,
): boolean {
  const collection = dvh.model?.collections.find((c) => c.id === collectionId)
  if (!collection) return false
  const { first, last } = selectedBlocks(editor)
  const single = first === last ? editor.state.doc.child(first) : null
  const kind = single?.type.name === 'docTable' ? 'repeatrows' : 'repeat'
  const tag = `dvh:${kind}:${collection.id}${conditionId ? `:${conditionId}` : ''}`
  return wrapSelection(editor, tag, collection.name)
}

/** Inserts a column value control at the caret (shows «Title» until generated). */
export function insertColumnControl(
  editor: Editor,
  dvh: DvhDocsState,
  collectionId: string,
  columnId: string,
): boolean {
  const collection = dvh.model?.collections.find((c) => c.id === collectionId)
  const column = collection?.columns.find((c) => c.id === columnId)
  const mark = editor.schema.marks.dvhField
  if (!collection || !column || !mark) return false
  const sdtPr = `<w:sdtPr><w:alias w:val="${escapeXml(`${collection.name}.${column.title}`)}"/><w:tag w:val="dvh:c:${escapeXml(column.id)}"/></w:sdtPr>`
  const node = editor.schema.text(`«${column.title}»`, [mark.create({ sdtPr })])
  editor.view.dispatch(editor.state.tr.replaceSelectionWith(node, false).scrollIntoView())
  return true
}

/**
 * Takes Smart Data imported from a QLCL workbook into the document: field
 * values (every occurrence follows), new fields, and collections (tables
 * showing them refresh). The import kept the ids of objects with the same
 * names, so occurrences and tables stay bound. One change set records it.
 */
export function mergeImportedData(editor: Editor, dvh: DvhDocsState, imported: DvhModel): void {
  const { model } = ensureModel(dvh)
  const before = structuredClone(model)
  for (const field of imported.fields) {
    const existing = model.fields.find((f) => f.id === field.id || f.name === field.name)
    if (!existing) model.fields.push({ ...field })
    else if (existing.value !== field.value) {
      applyFieldText(editor, dvh, existing.id, fieldText(field.value), 'Data.ImportQlcl', 'ui')
    }
  }
  for (const collection of imported.collections) {
    model.collections = [
      ...model.collections.filter((c) => c.id !== collection.id && c.name !== collection.name),
      structuredClone(collection),
    ]
    for (const table of model.tables) {
      if ('collectionId' in table.source && table.source.collectionId === collection.id) {
        refreshDocTable(editor, dvh, table.id, { force: true, source: 'ui' })
      }
    }
  }
  const changes = diffModels(before, model).filter((c) => c.path !== 'value')
  if (changes.length > 0) recordDvhDocsChange(dvh, 'Data.ImportQlcl', changes)
}

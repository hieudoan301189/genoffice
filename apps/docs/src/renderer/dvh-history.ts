/**
 * Object history in Docs (P5): restore one field, collection or table to the
 * state right after any change set, or undo one transaction (an AI run, a
 * workflow) by its txId — without rolling the whole file back. A restore is
 * itself a new change set (`source: restore`); nothing is removed from the
 * history.
 */
import type { Editor } from '@tiptap/core'
import {
  collectionData,
  dvhCollectionSchema,
  dvhTableSchema,
  fieldText,
  revertPlan,
  tableDefinition,
  valueAfter,
  type ChangeSet,
  type DvhCollection,
  type DvhTable,
  type HistoryWrite,
  type Scalar,
} from '@genoffice/dvh-model'
import {
  applyFieldText,
  historyEntries,
  recordDvhDocsChange,
  type DvhDocsState,
} from './dvh-smart-data'
import { refreshDocTable } from './dvh-tables'

/** What a write did, for the caller's report. */
export interface RestoreOutcome {
  readonly objectId: string
  readonly path: string
  readonly applied: boolean
  readonly reason?: string
}

/** All change sets of the document, oldest first (file history, then this session). */
export function allChangeSets(dvh: DvhDocsState | null): ChangeSet[] {
  return [...historyEntries(dvh)].reverse()
}

/** Re-renders every table that shows a collection (forced: the restore is the user's choice). */
function refreshTablesOf(editor: Editor, dvh: DvhDocsState, collectionId: string): void {
  for (const table of dvh.model?.tables ?? []) {
    if ('collectionId' in table.source && table.source.collectionId === collectionId) {
      refreshDocTable(editor, dvh, table.id, { force: true, source: 'restore' })
    }
  }
}

/**
 * Puts one object/path back to `value`. Supported: field values, collection
 * data, table definitions (embedded rows included), link update modes.
 */
export function applyHistoryWrite(
  editor: Editor,
  dvh: DvhDocsState,
  write: HistoryWrite,
  action: string,
): RestoreOutcome {
  const model = dvh.model
  const base = { objectId: write.objectId, path: write.path }
  if (!model) return { ...base, applied: false, reason: 'no Smart Data in this document' }

  const field = model.fields.find((f) => f.id === write.objectId)
  if (field && write.path === 'value') {
    applyFieldText(
      editor,
      dvh,
      field.id,
      fieldText((write.value ?? null) as Scalar),
      action,
      'restore',
    )
    return { ...base, applied: true }
  }

  const collection = model.collections.find((c) => c.id === write.objectId)
  if (collection && write.path === 'data') {
    if (write.value === null) {
      return { ...base, applied: false, reason: 'the collection did not exist then' }
    }
    const data = write.value as Pick<DvhCollection, 'columns' | 'rows'>
    const next = dvhCollectionSchema.parse({ ...collection, ...data })
    const before = collectionData(collection)
    model.collections = model.collections.map((c) => (c.id === next.id ? next : c))
    recordDvhDocsChange(
      dvh,
      action,
      [{ objectId: next.id, path: 'data', before, after: collectionData(next) }],
      'restore',
    )
    refreshTablesOf(editor, dvh, next.id)
    return { ...base, applied: true }
  }

  const table = model.tables.find((t) => t.id === write.objectId)
  if (table && write.path === 'definition') {
    if (write.value === null)
      return { ...base, applied: false, reason: 'the table did not exist then' }
    const next: DvhTable = dvhTableSchema.parse({
      ...(write.value as object),
      id: table.id,
      ...(table.lastRender ? { lastRender: table.lastRender } : {}),
    })
    const before = tableDefinition(table)
    model.tables = model.tables.map((t) => (t.id === table.id ? next : t))
    recordDvhDocsChange(
      dvh,
      action,
      [{ objectId: table.id, path: 'definition', before, after: tableDefinition(next) }],
      'restore',
    )
    refreshDocTable(editor, dvh, table.id, { force: true, source: 'restore' })
    return { ...base, applied: true }
  }

  const link = model.links.find((l) => l.id === write.objectId)
  if (link && write.path === 'update' && typeof write.value === 'string') {
    const before = link.update
    link.update = write.value as typeof link.update
    recordDvhDocsChange(dvh, action, [{ ...base, before, after: write.value }], 'restore')
    return { ...base, applied: true }
  }
  return { ...base, applied: false, reason: 'this change cannot be restored' }
}

/** Restores one object to its state right after a change set. */
export function restoreObject(
  editor: Editor,
  dvh: DvhDocsState,
  changeSetId: string,
  objectId: string,
): RestoreOutcome[] {
  const changeSet = allChangeSets(dvh).find((cs) => cs.id === changeSetId)
  if (!changeSet) throw new Error(`no change set ${changeSetId}`)
  const paths = [
    ...new Set(changeSet.changes.filter((c) => c.objectId === objectId).map((c) => c.path)),
  ]
  if (paths.length === 0) throw new Error(`change set ${changeSetId} did not touch ${objectId}`)
  return paths.map((path) =>
    applyHistoryWrite(
      editor,
      dvh,
      { objectId, path, value: valueAfter(changeSet, objectId, path).value },
      'History.RestoreObject',
    ),
  )
}

export interface RevertOutcome {
  readonly txId: string
  readonly results: readonly RestoreOutcome[]
  /** later edits of the same objects: the revert did not run (pass force to override) */
  readonly conflicts: readonly { objectId: string; path: string; changeSet: string }[]
}

/**
 * Undoes one transaction. When later change sets touched the same objects,
 * nothing runs unless `force` is set: the caller asks the user first.
 */
export function revertTransaction(
  editor: Editor,
  dvh: DvhDocsState,
  txId: string,
  options: { force?: boolean } = {},
): RevertOutcome {
  const plan = revertPlan(allChangeSets(dvh), txId)
  if (plan.writes.length === 0) throw new Error(`no change set with txId ${txId}`)
  if (plan.conflicts.length > 0 && !options.force) {
    return { txId, results: [], conflicts: plan.conflicts }
  }
  const results = plan.writes.map((w) =>
    applyHistoryWrite(editor, dvh, w, 'History.RevertTransaction'),
  )
  return { txId, results, conflicts: plan.conflicts }
}

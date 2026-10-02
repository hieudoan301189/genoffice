/**
 * Change sets (ADR D7) and the embedded history part (ADR D9):
 * `<dvh:history xmlns:dvh="urn:dvh-office:history:1">` holding one JSON change
 * set per line in CDATA (the zip already compresses it), plus the model hash at
 * the last DVH save so an edit made in Word/Excel can be detected on open.
 */

import { z } from 'zod'

import { decodeXml, escapeXml, HISTORY_NS } from './custom-xml'
import { dvhModelSchema, type DvhCollection, type DvhModel, type DvhTable } from './types'

export const changeSetSchema = z
  .object({
    id: z.string().min(1),
    txId: z.string().min(1),
    docId: z.string().min(1),
    at: z.string(),
    source: z.enum(['ui', 'ai', 'workflow', 'script', 'link', 'restore', 'external']),
    actor: z
      .object({ user: z.string().optional(), device: z.string().optional() })
      .strict()
      .optional(),
    action: z.string().min(1),
    changes: z.array(
      z
        .object({
          objectId: z.string().min(1),
          path: z.string(),
          before: z.unknown(),
          after: z.unknown(),
        })
        .strict(),
    ),
  })
  .strict()
export type ChangeSet = z.infer<typeof changeSetSchema>

/** The DVH model at one point of the history (P5): compaction anchors and restore points. */
export interface HistorySnapshot {
  readonly at: string
  /** the change set the snapshot was taken after (null: before the first one) */
  readonly afterChangeSet: string | null
  readonly model: DvhModel
}

export interface HistoryPart {
  readonly docId: string
  /** modelHash() of the model at the last DVH save */
  readonly modelHash?: string
  /**
   * the model at the last DVH save: when Word or Excel changes the file, the
   * next open diffs against it to record what changed outside DVH
   */
  readonly base?: DvhModel
  readonly snapshots?: readonly HistorySnapshot[]
  readonly changes: readonly ChangeSet[]
}

const cdata = (text: string) => `<![CDATA[${text.replace(/\]\]>/g, ']]]]><![CDATA[>')}]]>`

export function serializeHistoryXml(history: HistoryPart): string {
  const lines = history.changes.map((c) => JSON.stringify(c)).join('\n')
  const snapshots = (history.snapshots ?? []).map((s) => JSON.stringify(s)).join('\n')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<dvh:history xmlns:dvh="${HISTORY_NS}" docId="${escapeXml(history.docId)}"` +
    (history.modelHash ? ` modelHash="${history.modelHash}"` : '') +
    ` count="${history.changes.length}">` +
    `<dvh:changes>${cdata(lines)}</dvh:changes>` +
    (history.base ? `<dvh:base>${cdata(JSON.stringify(history.base))}</dvh:base>` : '') +
    (snapshots ? `<dvh:snapshots>${cdata(snapshots)}</dvh:snapshots>` : '') +
    '</dvh:history>'
  )
}

export function parseHistoryXml(xml: string): HistoryPart {
  const root = /<(?:\w+:)?history\b[^>]*>/.exec(xml)
  if (!root || !xml.includes(HISTORY_NS)) throw new Error('not a DVH history part')
  const attr = (name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(root[0])?.[1]
  const body = /<(?:\w+:)?changes\b[^>]*>([\s\S]*?)<\/(?:\w+:)?changes>/.exec(xml)?.[1] ?? ''
  const changes = decodeXml(body)
    .split('\n')
    .filter((line) => line.trim() !== '')
    .map((line) => changeSetSchema.parse(JSON.parse(line)))
  const modelHash = attr('modelHash')
  const element = (name: string) =>
    new RegExp(`<(?:\\w+:)?${name}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${name}>`).exec(xml)?.[1]
  // base and snapshots are optional extras: a damaged one is dropped, never the change sets
  let base: DvhModel | undefined
  try {
    const text = element('base')
    if (text) base = dvhModelSchema.parse(JSON.parse(decodeXml(text)))
  } catch {
    base = undefined
  }
  const snapshots: HistorySnapshot[] = []
  for (const line of decodeXml(element('snapshots') ?? '').split('\n')) {
    if (line.trim() === '') continue
    try {
      const raw = JSON.parse(line) as HistorySnapshot
      snapshots.push({ ...raw, model: dvhModelSchema.parse(raw.model) })
    } catch {
      // skip the damaged snapshot
    }
  }
  return {
    docId: decodeXml(attr('docId') ?? ''),
    ...(modelHash ? { modelHash } : {}),
    ...(base ? { base } : {}),
    ...(snapshots.length ? { snapshots } : {}),
    changes,
  }
}

/** Order-independent JSON for hashing. */
function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`
  if (value && typeof value === 'object') {
    return `{${Object.keys(value as Record<string, unknown>)
      .sort()
      .map((k) => `${JSON.stringify(k)}:${canonical((value as Record<string, unknown>)[k])}`)
      .join(',')}}`
  }
  return JSON.stringify(value) ?? 'null'
}

/** 64-bit FNV-1a of the canonical model, hex; synchronous so renderers can call it. */
export function modelHash(model: DvhModel): string {
  return contentHash(model)
}

/** 64-bit FNV-1a of any JSON value in canonical form (key order does not matter), hex. */
export function contentHash(value: unknown): string {
  let hash = 0xcbf29ce484222325n
  const prime = 0x100000001b3n
  for (const byte of new TextEncoder().encode(canonical(value))) {
    hash ^= BigInt(byte)
    hash = (hash * prime) & 0xffffffffffffffffn
  }
  return hash.toString(16).padStart(16, '0')
}

/** The data of a collection as history records it: columns and rows, nothing derived. */
export function collectionData(collection: DvhCollection): Pick<DvhCollection, 'columns' | 'rows'> {
  return { columns: collection.columns, rows: collection.rows }
}

/** A table's settings as history records them (what it rendered last is not history). */
export function tableDefinition(table: DvhTable): Omit<DvhTable, 'lastRender'> {
  const { lastRender: _lastRender, ...definition } = table
  return definition
}

/**
 * Differences between two models, as change-set entries: field values and
 * names, collection data (`data`, before/after columns and rows), table
 * settings (`definition`), link sources and modes, objects added or removed.
 */
export function diffModels(before: DvhModel, after: DvhModel): ChangeSet['changes'] {
  const changes: ChangeSet['changes'] = []
  const old = new Map(before.fields.map((f) => [f.id, f]))
  for (const field of after.fields) {
    const prev = old.get(field.id)
    if (!prev) changes.push({ objectId: field.id, path: '', before: null, after: field })
    else {
      if (prev.value !== field.value)
        changes.push({ objectId: field.id, path: 'value', before: prev.value, after: field.value })
      if (prev.name !== field.name)
        changes.push({ objectId: field.id, path: 'name', before: prev.name, after: field.name })
    }
    old.delete(field.id)
  }
  for (const removed of old.values())
    changes.push({ objectId: removed.id, path: '', before: removed, after: null })
  const byId = <T extends { id: string }>(list: readonly T[]) => new Map(list.map((x) => [x.id, x]))
  const oldCollections = byId(before.collections)
  for (const collection of after.collections) {
    const prev = oldCollections.get(collection.id)
    oldCollections.delete(collection.id)
    const next = collectionData(collection)
    if (!prev) changes.push({ objectId: collection.id, path: 'data', before: null, after: next })
    else if (contentHash(collectionData(prev)) !== contentHash(next))
      changes.push({
        objectId: collection.id,
        path: 'data',
        before: collectionData(prev),
        after: next,
      })
  }
  for (const removed of oldCollections.values())
    changes.push({
      objectId: removed.id,
      path: 'data',
      before: collectionData(removed),
      after: null,
    })
  const oldTables = byId(before.tables)
  for (const table of after.tables) {
    const prev = oldTables.get(table.id)
    oldTables.delete(table.id)
    const next = tableDefinition(table)
    if (!prev) changes.push({ objectId: table.id, path: 'definition', before: null, after: next })
    else if (contentHash(tableDefinition(prev)) !== contentHash(next))
      changes.push({
        objectId: table.id,
        path: 'definition',
        before: tableDefinition(prev),
        after: next,
      })
  }
  for (const removed of oldTables.values())
    changes.push({
      objectId: removed.id,
      path: 'definition',
      before: tableDefinition(removed),
      after: null,
    })
  const oldLinks = byId(before.links)
  for (const link of after.links) {
    const prev = oldLinks.get(link.id)
    oldLinks.delete(link.id)
    if (!prev) changes.push({ objectId: link.id, path: '', before: null, after: link.source })
    else {
      if (contentHash(prev.source) !== contentHash(link.source))
        changes.push({ objectId: link.id, path: 'source', before: prev.source, after: link.source })
      if (prev.update !== link.update)
        changes.push({ objectId: link.id, path: 'update', before: prev.update, after: link.update })
    }
  }
  for (const removed of oldLinks.values())
    changes.push({ objectId: removed.id, path: '', before: removed.source, after: null })
  return changes
}

// ---------------- P5: object history ----------------

/** P0 decision (ADR D9): 5 MB of history JSON adds ~0.4 MB to a file and opens fast in Office. */
export const HISTORY_SIZE_LIMIT = 5 * 1024 * 1024
/** a snapshot of the model is kept every this many change sets */
export const SNAPSHOT_EVERY = 100
/** compaction keeps the change sets after this many most recent snapshots */
export const KEEP_SNAPSHOTS = 2

/** Bytes of history JSON (change sets, base, snapshots) as the part stores it. */
export function historySizeBytes(history: HistoryPart): number {
  return new TextEncoder().encode(serializeHistoryXml(history)).length
}

/**
 * The history part to write on save: the session's change sets appended, the
 * model at this save as the base for external-edit detection, and a snapshot
 * when SNAPSHOT_EVERY change sets passed since the last one.
 */
export function historyForSave(
  previous: HistoryPart | null,
  pending: readonly ChangeSet[],
  model: DvhModel,
  now: () => string = () => new Date().toISOString(),
): HistoryPart {
  const changes = [...(previous?.changes ?? []), ...pending]
  const snapshots = [...(previous?.snapshots ?? [])]
  const last = snapshots.at(-1)
  const sinceLast = last
    ? changes.length - 1 - changes.findIndex((c) => c.id === last.afterChangeSet)
    : changes.length
  if (changes.length > 0 && sinceLast >= SNAPSHOT_EVERY) {
    snapshots.push({ at: now(), afterChangeSet: changes.at(-1)!.id, model: structuredClone(model) })
  }
  return {
    docId: model.docId,
    modelHash: modelHash(model),
    base: structuredClone(model),
    ...(snapshots.length ? { snapshots } : {}),
    changes,
  }
}

/**
 * Compaction: every change set older than the KEEP_SNAPSHOTS-th most recent
 * snapshot becomes one `History.Compact` change set holding, per object and
 * path, the first `before` and the last `after`; older snapshots are dropped.
 * Returns the history unchanged when there is nothing to compact.
 */
export function compactHistory(history: HistoryPart, keep: number = KEEP_SNAPSHOTS): HistoryPart {
  const snapshots = history.snapshots ?? []
  const anchor = snapshots.length >= keep ? snapshots[snapshots.length - keep] : undefined
  const cut = anchor ? history.changes.findIndex((c) => c.id === anchor.afterChangeSet) + 1 : 0
  return compactBefore(history, cut)
}

/** Merges the change sets before index `cut` into one; snapshots inside that span go. */
function compactBefore(history: HistoryPart, cut: number): HistoryPart {
  if (cut <= 1) return history
  const old = history.changes.slice(0, cut)
  const merged = new Map<string, ChangeSet['changes'][number]>()
  for (const cs of old) {
    for (const change of cs.changes) {
      const key = `${change.objectId}\u0000${change.path}`
      const prev = merged.get(key)
      merged.set(key, prev ? { ...prev, after: change.after } : { ...change })
    }
  }
  const last = old.at(-1)!
  const compacted: ChangeSet = {
    id: `cs_compact_${last.id}`,
    txId: `tx_compact_${last.id}`,
    docId: history.docId,
    at: last.at,
    source: 'restore',
    action: 'History.Compact',
    changes: [...merged.values()].filter((c) => contentHash(c.before) !== contentHash(c.after)),
  }
  const kept = new Set(history.changes.slice(cut - 1).map((c) => c.id))
  const snapshots = (history.snapshots ?? []).filter(
    (s) => s.afterChangeSet !== null && kept.has(s.afterChangeSet),
  )
  return {
    ...history,
    ...(snapshots.length ? { snapshots } : { snapshots: [] }),
    changes: [compacted, ...history.changes.slice(cut)],
  }
}

/**
 * Compaction on request (the user saw the size warning): everything but the
 * `keepRecent` newest change sets becomes one `History.Compact` change set;
 * snapshots of the compacted span are dropped.
 */
export function compactHistoryKeeping(history: HistoryPart, keepRecent: number): HistoryPart {
  return compactBefore(history, history.changes.length - Math.max(0, keepRecent))
}

/**
 * A file saved by Word or Excel after its last DVH save carries another model:
 * the difference becomes one change set from `external`, so the history has
 * no silent gap. Null when the model is what DVH saved.
 */
export function externalChangeSet(
  history: HistoryPart | null,
  current: DvhModel,
  now: () => string = () => new Date().toISOString(),
): ChangeSet | null {
  if (!history?.modelHash || history.modelHash === modelHash(current)) return null
  const changes = history.base
    ? diffModels(history.base, current)
    : [
        {
          objectId: current.docId,
          path: 'model',
          before: history.modelHash,
          after: modelHash(current),
        },
      ]
  if (changes.length === 0) return null
  const stamp = now()
  return {
    id: `cs_ext_${modelHash(current)}`,
    txId: `tx_ext_${modelHash(current)}`,
    docId: current.docId,
    at: stamp,
    source: 'external',
    action: 'File.EditedOutside',
    changes,
  }
}

/** The change sets that touched an object, oldest first. */
export function objectHistory(changes: readonly ChangeSet[], objectId: string): ChangeSet[] {
  return changes.filter((cs) => cs.changes.some((c) => c.objectId === objectId))
}

/** Keeps the change sets made at or after `from` (ISO time): "keep from a point" on export. */
export function historyFrom(history: HistoryPart, from: string): HistoryPart {
  return {
    ...history,
    snapshots: (history.snapshots ?? []).filter((s) => s.at >= from),
    changes: history.changes.filter((c) => c.at >= from),
  }
}

/** One write a revert or a restore needs: put `path` of `objectId` back to `value`. */
export interface HistoryWrite {
  readonly objectId: string
  readonly path: string
  readonly value: unknown
}

export interface RevertPlan {
  readonly writes: readonly HistoryWrite[]
  /** later change sets that touched the same object and path: asked about before reverting */
  readonly conflicts: readonly {
    readonly objectId: string
    readonly path: string
    readonly changeSet: string
  }[]
}

/**
 * Undoing a transaction (P5): every object and path it changed goes back to
 * its value before the transaction. A change made later to the same object
 * and path is a conflict; the caller asks before running the plan.
 */
export function revertPlan(changes: readonly ChangeSet[], txId: string): RevertPlan {
  const first = changes.findIndex((cs) => cs.txId === txId)
  if (first < 0) return { writes: [], conflicts: [] }
  const inTx = changes.filter((cs) => cs.txId === txId)
  const lastIndex = changes.map((cs) => cs.txId).lastIndexOf(txId)
  const writes = new Map<string, HistoryWrite>()
  for (const cs of inTx) {
    for (const c of cs.changes) {
      const key = `${c.objectId}\u0000${c.path}`
      // the first before of the transaction is where the object was
      if (!writes.has(key)) writes.set(key, { objectId: c.objectId, path: c.path, value: c.before })
    }
  }
  const conflicts: { objectId: string; path: string; changeSet: string }[] = []
  for (const cs of changes.slice(lastIndex + 1)) {
    if (cs.txId === txId) continue
    for (const c of cs.changes) {
      if (writes.has(`${c.objectId}\u0000${c.path}`))
        conflicts.push({ objectId: c.objectId, path: c.path, changeSet: cs.id })
    }
  }
  return { writes: [...writes.values()], conflicts }
}

/** The value of `objectId`/`path` right after a change set (a restore point). */
export function valueAfter(
  changeSet: ChangeSet,
  objectId: string,
  path: string,
): { found: boolean; value: unknown } {
  const hits = changeSet.changes.filter((c) => c.objectId === objectId && c.path === path)
  return hits.length
    ? { found: true, value: hits.at(-1)!.after }
    : { found: false, value: undefined }
}

/** Cell-level differences between two versions of a collection's data, for the history panel. */
export function cellDiff(
  before: Pick<DvhCollection, 'columns' | 'rows'> | null,
  after: Pick<DvhCollection, 'columns' | 'rows'> | null,
): {
  readonly columnsAdded: readonly string[]
  readonly columnsRemoved: readonly string[]
  readonly rowsBefore: number
  readonly rowsAfter: number
  readonly cells: readonly { row: number; column: string; before: unknown; after: unknown }[]
  readonly styleChanged: boolean
} {
  const bCols = before?.columns ?? []
  const aCols = after?.columns ?? []
  const bIds = new Set(bCols.map((c) => c.id))
  const aIds = new Set(aCols.map((c) => c.id))
  const cells: { row: number; column: string; before: unknown; after: unknown }[] = []
  const rows = Math.max(before?.rows.length ?? 0, after?.rows.length ?? 0)
  for (const column of aCols) {
    if (!bIds.has(column.id)) continue
    const bi = bCols.findIndex((c) => c.id === column.id)
    const ai = aCols.findIndex((c) => c.id === column.id)
    for (let r = 0; r < rows; r++) {
      const was = before?.rows[r]?.[bi] ?? null
      const now = after?.rows[r]?.[ai] ?? null
      if (was !== now) cells.push({ row: r, column: column.title, before: was, after: now })
    }
  }
  const style = (cols: typeof bCols) =>
    contentHash(cols.map((c) => [c.id, c.style ?? null, c.headerStyle ?? null, c.numFmt ?? null]))
  return {
    columnsAdded: aCols.filter((c) => !bIds.has(c.id)).map((c) => c.title),
    columnsRemoved: bCols.filter((c) => !aIds.has(c.id)).map((c) => c.title),
    rowsBefore: before?.rows.length ?? 0,
    rowsAfter: after?.rows.length ?? 0,
    cells,
    styleChanged:
      style(bCols.filter((c) => aIds.has(c.id))) !== style(aCols.filter((c) => bIds.has(c.id))),
  }
}

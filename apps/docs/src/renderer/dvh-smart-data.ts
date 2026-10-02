/**
 * DVH Smart Data in Docs (P1, ADR D1/D2/D7/D9).
 *
 * The document keeps its own model part (fields copied from a linked
 * workbook, plus the link) and Smart Field content controls bound to it with
 * w:dataBinding. Text typed inside a field flows back into the model on save
 * (two-way, as Word does), so Word shows the same value when it re-reads the
 * store. "Update from source" re-reads the workbook's model part from disk.
 */
import type { Editor } from '@tiptap/core'
import type { Mark, Node as PmNode, Schema } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import {
  dvhFieldId,
  readCustomXmlPart,
  type CustomXmlPartWrite,
  type ParsedDocFull,
} from '@genoffice/docx-engine'
import {
  collectionData,
  compactHistoryKeeping,
  emptyModel,
  externalChangeSet,
  fieldBaseOf,
  historyForSave,
  historySizeBytes,
  HISTORY_SIZE_LIMIT,
  fieldSdtPrXml,
  fieldText,
  fieldValueFromText,
  contentHash,
  HISTORY_NS,
  MODEL_NS,
  modelHash,
  newDvhId,
  newStoreItemId,
  parseHistoryXml,
  parseModelXml,
  serializeHistoryXml,
  serializeModelXml,
  type ChangeSet,
  type DvhField,
  type DvhLink,
  type DvhModel,
  type FieldBase,
  type HistoryPart,
  type Scalar,
  threeWay,
} from '@genoffice/dvh-model'

export interface DvhDocsState {
  model: DvhModel | null
  modelStoreItemId: string | null
  history: HistoryPart | null
  historyStoreItemId: string | null
  /** change sets made this session, appended to the history part on save */
  pending: ChangeSet[]
  /**
   * change sets a previous session made and never saved (the app quit or
   * crashed): offered back to the user, never merged silently
   */
  recovered?: ChangeSet[]
}

/** A DvhDocsState whose model exists (after ensureModel). */
export type DvhDocsReady = DvhDocsState & { model: DvhModel; modelStoreItemId: string }

const states = new WeakMap<ParsedDocFull, DvhDocsState>()
let active: DvhDocsState | null = null
let onModelChange: (() => void) | null = null

export function dvhDocsStateOf(parsed: ParsedDocFull | null | undefined): DvhDocsState | null {
  return parsed ? (states.get(parsed) ?? null) : null
}

/** The open document's Smart Data (one document per Docs window); ops and actions act on it. */
export function activeDvhDocs(): DvhDocsState | null {
  return active
}

export function setActiveDvhDocs(state: DvhDocsState | null): void {
  active = state
}

/** Called when a model edit happens outside the editor content (the document becomes dirty). */
export function setDvhModelChangeListener(listener: (() => void) | null): void {
  onModelChange = listener
}

/** Reads the document's model/history parts (once per parsed document). */
export async function loadDvhDocs(parsed: ParsedDocFull): Promise<DvhDocsState> {
  const known = states.get(parsed)
  if (known) return known
  const bytes = parsed.internal.originalBytes
  const [modelPart, historyPart] = await Promise.all([
    readCustomXmlPart(bytes, MODEL_NS),
    readCustomXmlPart(bytes, HISTORY_NS),
  ])
  const state: DvhDocsState = {
    model: null,
    modelStoreItemId: modelPart?.storeItemId ?? null,
    history: null,
    historyStoreItemId: historyPart?.storeItemId ?? null,
    pending: [],
  }
  try {
    state.model = modelPart ? parseModelXml(modelPart.xml) : null
  } catch (error) {
    console.warn('DVH model part ignored:', error)
  }
  try {
    state.history = historyPart ? parseHistoryXml(historyPart.xml) : null
  } catch (error) {
    console.warn('DVH history part ignored:', error)
  }
  // P5: a model Word changed since the last DVH save is recorded, not silently adopted
  if (state.model) {
    const external = externalChangeSet(state.history, state.model)
    if (external) state.pending.push(external)
    await readBufferedHistory(state)
  }
  states.set(parsed, state)
  return state
}

/** Change sets a previous session left unsaved; a buffer fully in the file is cleared. */
async function readBufferedHistory(state: DvhDocsState): Promise<void> {
  const docId = state.model?.docId
  if (!docId || typeof window === 'undefined' || !window.desktop?.dvhReadHistoryBuffer) return
  try {
    const buffered = await window.desktop.dvhReadHistoryBuffer(docId)
    const saved = new Set((state.history?.changes ?? []).map((c) => c.id))
    const left = buffered.filter((c) => !saved.has(c.id))
    if (left.length > 0) state.recovered = left
    else if (buffered.length > 0) void window.desktop.dvhBufferHistory?.(docId, [])
  } catch {
    // the buffer is a safety net; reading it never blocks opening the document
  }
}

let bufferTimer: ReturnType<typeof setTimeout> | undefined
/** Mirrors the session's change sets to the main process' buffer (debounced). */
function scheduleHistoryBuffer(dvh: DvhDocsState): void {
  if (typeof window === 'undefined' || !window.desktop?.dvhBufferHistory) return
  clearTimeout(bufferTimer)
  bufferTimer = setTimeout(() => {
    const docId = dvh.model?.docId
    if (docId) void window.desktop.dvhBufferHistory(docId, dvh.pending).catch(() => {})
  }, 1000)
}

/** Takes the change sets of a previous session into this one's history (the user agreed). */
export function adoptRecoveredHistory(dvh: DvhDocsState): number {
  const recovered = dvh.recovered ?? []
  dvh.pending.unshift(...recovered)
  dvh.recovered = undefined
  if (recovered.length > 0) onModelChange?.()
  return recovered.length
}

/** Drops the change sets of a previous session (the user declined). */
export function discardRecoveredHistory(dvh: DvhDocsState): void {
  dvh.recovered = undefined
  const docId = dvh.model?.docId
  if (docId) void window.desktop?.dvhBufferHistory?.(docId, dvh.pending)
}

/** History size against the P0 limit (bytes of the part as it would be saved). */
export function historyStatus(dvh: DvhDocsState | null): { bytes: number; limit: number } {
  if (!dvh?.model) return { bytes: 0, limit: HISTORY_SIZE_LIMIT }
  return {
    bytes: historySizeBytes(historyForSave(dvh.history, dvh.pending, dvh.model)),
    limit: HISTORY_SIZE_LIMIT,
  }
}

/** states whose history was rewritten (compacted) and must be saved even with no pending change */
const historyRewritten = new WeakSet<DvhDocsState>()

/** Compacts the history on request, keeping the newest change sets as they are. */
export function compactDvhHistory(dvh: DvhDocsState, keepRecent = 200): void {
  if (!dvh.model) return
  const merged: HistoryPart = {
    ...(dvh.history ?? { docId: dvh.model.docId, changes: [] }),
    changes: [...(dvh.history?.changes ?? []), ...dvh.pending],
  }
  dvh.history = compactHistoryKeeping(merged, keepRecent)
  dvh.pending = []
  historyRewritten.add(dvh)
  onModelChange?.()
}

export function ensureModel(dvh: DvhDocsState): DvhDocsReady {
  dvh.model ??= emptyModel(newDvhId('doc'))
  dvh.modelStoreItemId ??= newStoreItemId()
  return dvh as DvhDocsReady
}

/**
 * Appends a change set. `external` marks edits made outside the editor content
 * (links, data actions), which the editor's own dirty tracking cannot see.
 */
/** the action run in progress (P7): its change sets share its txId and caller */
let ambient: { txId: string; source: ChangeSet['source'] } | null = null

const SOURCE_OF_CALLER: Record<string, ChangeSet['source']> = {
  ai: 'ai',
  workflow: 'workflow',
  script: 'script',
}

/**
 * Runs an action with its txId and caller ambient (ActionRegistry `around`):
 * what this module records meanwhile belongs to that run, so an AI operation
 * or a workflow is undone as one transaction (History.RevertTransaction).
 */
export async function withDvhTransaction<T>(
  txId: string,
  caller: string,
  run: () => Promise<T>,
): Promise<T> {
  const previous = ambient
  ambient = { txId, source: SOURCE_OF_CALLER[caller] ?? 'ui' }
  try {
    return await run()
  } finally {
    ambient = previous
  }
}

function record(
  dvh: DvhDocsState,
  action: string,
  changes: ChangeSet['changes'],
  source: ChangeSet['source'],
  external: boolean,
): void {
  if (!dvh.model || changes.length === 0) return
  dvh.pending.push({
    id: newDvhId('r').replace(/^r_/, 'cs_'),
    txId: ambient?.txId ?? newDvhId('r').replace(/^r_/, 'tx_'),
    docId: dvh.model.docId,
    at: new Date().toISOString(),
    // a run's caller replaces the generic 'ui'; link/restore/external keep their meaning
    source: ambient && source === 'ui' ? ambient.source : source,
    action,
    changes,
  })
  scheduleHistoryBuffer(dvh)
  if (external) onModelChange?.()
}

/** automatic link updates of this session that a following automatic update may extend */
const coalescible = new WeakSet<ChangeSet>()

/**
 * Automatic link updates arrive in bursts (every edit in Sheets): consecutive
 * ones merge into one change set — first `before`, last `after` per object
 * and path — so the history keeps one entry per burst instead of hundreds.
 */
function recordCoalesced(dvh: DvhDocsState, action: string, changes: ChangeSet['changes']): void {
  const last = dvh.pending.at(-1)
  if (!last || !coalescible.has(last) || last.action !== action || changes.length === 0) {
    record(dvh, action, changes, 'link', true)
    const added = dvh.pending.at(-1)
    if (added && added !== last) coalescible.add(added)
    return
  }
  const merged = [...last.changes]
  for (const change of changes) {
    const i = merged.findIndex((c) => c.objectId === change.objectId && c.path === change.path)
    if (i >= 0) merged[i] = { ...merged[i]!, after: change.after }
    else merged.push(change)
  }
  const next: ChangeSet = { ...last, at: new Date().toISOString(), changes: merged }
  dvh.pending[dvh.pending.length - 1] = next
  coalescible.add(next)
  scheduleHistoryBuffer(dvh)
  onModelChange?.()
}

/** A field by id or by name (`Project.Name`). */
export function findField(dvh: DvhDocsState | null, ref: string): DvhField | undefined {
  return (
    dvh?.model?.fields.find((f) => f.id === ref) ?? dvh?.model?.fields.find((f) => f.name === ref)
  )
}

export interface FieldOccurrence {
  readonly fieldId: string
  readonly sdtPr: string
  readonly from: number
  readonly to: number
  readonly text: string
}

/** Every Smart Field in the document; adjacent runs of one control merge. */
export function fieldOccurrences(doc: PmNode): FieldOccurrence[] {
  const out: FieldOccurrence[] = []
  doc.descendants((node, pos) => {
    if (!node.isText) return
    const mark = node.marks.find((m) => m.type.name === 'dvhField')
    const sdtPr = mark ? String(mark.attrs.sdtPr ?? '') : ''
    const fieldId = sdtPr ? dvhFieldId(sdtPr) : null
    if (!fieldId) return
    const last = out[out.length - 1]
    if (last && last.sdtPr === sdtPr && last.to === pos) {
      out[out.length - 1] = {
        ...last,
        to: pos + node.nodeSize,
        text: last.text + (node.text ?? ''),
      }
    } else {
      out.push({ fieldId, sdtPr, from: pos, to: pos + node.nodeSize, text: node.text ?? '' })
    }
  })
  return out
}

function usedSdtIds(doc: PmNode): Set<number> {
  const ids = new Set<number>()
  for (const occ of fieldOccurrences(doc)) {
    const m = /<w:id w:val="(-?\d+)"/.exec(occ.sdtPr)
    if (m) ids.add(Number(m[1]))
  }
  return ids
}

/**
 * The text node for one new occurrence of a field: its current value inside a
 * bound content control with a w:id unique in `doc`. An empty value shows the
 * field name, as a text node cannot be empty.
 */
export function smartFieldNode(
  schema: Schema,
  doc: PmNode,
  dvh: DvhDocsReady,
  field: DvhField,
  marks: readonly Mark[] = [],
): PmNode {
  const used = usedSdtIds(doc)
  let sdtId = 0
  while (sdtId === 0 || used.has(sdtId)) sdtId = 1 + Math.floor(Math.random() * 2_000_000_000)
  const sdtPr = fieldSdtPrXml({
    fieldId: field.id,
    alias: field.name,
    sdtId,
    storeItemId: dvh.modelStoreItemId,
  })
  return schema.text(fieldText(field.value) || field.name, [
    ...marks,
    schema.marks.dvhField!.create({ sdtPr }),
  ])
}

/** Adds the field to the model if it is new and records the insertion. */
export function noteFieldInserted(
  dvh: DvhDocsState,
  field: DvhField,
  source: ChangeSet['source'],
): void {
  const ready = ensureModel(dvh)
  if (!ready.model.fields.some((f) => f.id === field.id)) ready.model.fields.push({ ...field })
  record(
    dvh,
    'Document.InsertField',
    [{ objectId: field.id, path: 'occurrence', before: null, after: field.name }],
    source,
    false,
  )
}

/** Inserts a Smart Field control at the caret, showing the field's current value. */
export function insertSmartField(editor: Editor, dvh: DvhDocsState, field: DvhField): void {
  const ready = ensureModel(dvh)
  const { state } = editor
  const node = smartFieldNode(state.schema, state.doc, ready, field)
  editor.view.dispatch(state.tr.replaceSelectionWith(node, false).scrollIntoView())
  editor.commands.focus()
  noteFieldInserted(dvh, field, 'ui')
}

/**
 * Replaces the text of every occurrence of the given fields inside `tr`,
 * keeping each occurrence's marks; returns how many occurrences changed.
 */
export function replaceFieldTexts(tr: Transaction, texts: ReadonlyMap<string, string>): number {
  const doc = tr.doc
  const occurrences = fieldOccurrences(doc).filter(
    (occ) => texts.has(occ.fieldId) && texts.get(occ.fieldId) !== occ.text,
  )
  // back to front, so earlier positions stay valid
  for (const occ of [...occurrences].reverse()) {
    // the first node's own marks: resolve().marks() drops non-inclusive marks at a node's end
    const marks = doc.nodeAt(occ.from)?.marks ?? []
    const text = texts.get(occ.fieldId)!
    if (text === '') tr.delete(occ.from, occ.to)
    else tr.replaceWith(occ.from, occ.to, doc.type.schema.text(text, marks))
  }
  return occurrences.length
}

/** Sets a field's model value and records it (the occurrences are the caller's job). */
export function setFieldValue(
  dvh: DvhDocsState,
  field: DvhField,
  value: Scalar,
  action: string,
  source: ChangeSet['source'],
): void {
  if (value === field.value) return
  const before = field.value
  field.value = value
  record(dvh, action, [{ objectId: field.id, path: 'value', before, after: value }], source, true)
}

/** Two-way binding: field text edited in the document becomes the model value. */
export function syncModelFromEditor(editor: Editor, dvh: DvhDocsState): void {
  if (!dvh.model) return
  const firstText = new Map<string, string>()
  for (const occ of fieldOccurrences(editor.state.doc)) {
    if (!firstText.has(occ.fieldId)) firstText.set(occ.fieldId, occ.text)
  }
  const changes: ChangeSet['changes'] = []
  for (const field of dvh.model.fields) {
    const text = firstText.get(field.id)
    if (text === undefined || text === fieldText(field.value)) continue
    const value = fieldValueFromText(field.type, text)
    changes.push({ objectId: field.id, path: 'value', before: field.value, after: value })
    field.value = value
  }
  // typed in the page: the editor already marked the document dirty
  record(dvh, 'Document.EditField', changes, 'ui', false)
}

/** A workbook's model as read from disk, with where it was found. */
export interface DvhSource {
  readonly path: string
  readonly model: DvhModel
}

const folderOf = (p: string) => p.slice(0, Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/')) + 1)

/** Relative path from the document's folder when the workbook is inside it, else undefined. */
function relativeFromDocument(sourcePath: string, docPath: string | null): string | undefined {
  if (!docPath) return undefined
  const docDir = folderOf(docPath)
  return sourcePath.toLowerCase().startsWith(docDir.toLowerCase())
    ? sourcePath.slice(docDir.length)
    : undefined
}

/** Links a workbook: its fields join the document model (same ids). */
export function linkSource(dvh: DvhDocsState, source: DvhSource, docPath: string | null): DvhLink {
  const { model } = ensureModel(dvh)
  for (const field of source.model.fields) {
    const existing = model.fields.find((f) => f.id === field.id)
    if (existing)
      Object.assign(existing, { name: field.name, type: field.type, value: field.value })
    else model.fields.push({ ...field, access: 'readwrite' })
  }
  for (const collection of source.model.collections) {
    model.collections = [
      ...model.collections.filter((c) => c.id !== collection.id),
      structuredClone(collection),
    ]
  }
  const relPath = relativeFromDocument(source.path, docPath)
  const link: DvhLink = {
    id: model.links.find((l) => l.source.docId === source.model.docId)?.id ?? newDvhId('l'),
    source: {
      docId: source.model.docId,
      ...(relPath ? { relPath } : {}),
      path: source.path,
      objectId: '*',
    },
    targets: [
      ...source.model.fields.map((f) => f.id),
      ...source.model.collections.map((c) => c.id),
    ],
    update: 'manual',
    lastSync: {
      revision: 0,
      hash: modelHash(source.model),
      at: new Date().toISOString(),
      fields: Object.fromEntries(source.model.fields.map((f) => [f.id, fieldBaseOf(f)])),
    },
  }
  model.links = [...model.links.filter((l) => l.id !== link.id), link]
  record(
    dvh,
    'Link.Create',
    [{ objectId: link.id, path: '', before: null, after: link }],
    'link',
    true,
  )
  return link
}

/**
 * Points a link at the workbook's new place (found by docId after a rename or
 * a move, or chosen with "Change source"). `docId` changes only when the user
 * picked a copy that carries every object the link reads.
 */
export function relocateLink(
  dvh: DvhDocsState,
  link: DvhLink,
  path: string,
  docPath: string | null,
  docId: string = link.source.docId,
): void {
  const before = { docId: link.source.docId, path: link.source.path, relPath: link.source.relPath }
  const relPath = relativeFromDocument(path, docPath)
  link.source = {
    docId,
    ...(relPath ? { relPath } : {}),
    path,
    objectId: link.source.objectId,
  }
  if (before.path === path && before.relPath === relPath && before.docId === docId) return
  record(
    dvh,
    'Link.Relocate',
    [{ objectId: link.id, path: 'source', before, after: { docId, path, relPath } }],
    'link',
    true,
  )
}

/** Candidate paths for a link's workbook: next to the document first, then the last known path. */
export function linkSourcePaths(link: DvhLink, docPath: string | null): string[] {
  const paths: string[] = []
  if (link.source.relPath && docPath) paths.push(folderOf(docPath) + link.source.relPath)
  if (link.source.path) paths.push(link.source.path)
  return [...new Set(paths)]
}

/** True when the workbook changed since the link last pulled from it (stale values). */
export function linkIsStale(link: DvhLink, source: DvhSource): boolean {
  return link.lastSync?.hash !== modelHash(source.model)
}

/** A field changed on both sides since the link last synced, to different texts. */
export interface FieldConflict {
  readonly fieldId: string
  /** the document's text */
  readonly local: string
  /** the source as it is now */
  readonly source: FieldBase
  readonly base: FieldBase | undefined
}

export interface SourceUpdate {
  /** field occurrences whose text changed */
  readonly fields: number
  /** collections whose columns or rows changed (their tables need a refresh) */
  readonly collections: readonly string[]
  /** fields edited in the document only: kept, waiting to be written back */
  readonly localEdits: readonly string[]
  /** fields edited on both sides: kept as they are here until the user decides */
  readonly conflicts: readonly FieldConflict[]
}

/** The text a field shows in the document: its first occurrence, else its model value. */
export function localFieldTexts(doc: PmNode, dvh: DvhDocsState): Map<string, string> {
  const texts = new Map<string, string>()
  for (const occ of fieldOccurrences(doc)) {
    if (!texts.has(occ.fieldId)) texts.set(occ.fieldId, occ.text)
  }
  for (const field of dvh.model?.fields ?? []) {
    if (!texts.has(field.id)) texts.set(field.id, fieldText(field.value))
  }
  return texts
}

/** The base a link last agreed on for one field (undefined for links made before P3). */
export function linkFieldBase(link: DvhLink, fieldId: string): FieldBase | undefined {
  return link.lastSync?.fields?.[fieldId]
}

/** Moves a link's base for some fields (after a pull, a write-back or a resolved conflict). */
export function setLinkFieldBases(link: DvhLink, bases: ReadonlyMap<string, FieldBase>): void {
  if (bases.size === 0) return
  link.lastSync = {
    revision: link.lastSync?.revision ?? 0,
    hash: link.lastSync?.hash ?? '',
    at: link.lastSync?.at ?? new Date().toISOString(),
    fields: { ...(link.lastSync?.fields ?? {}), ...Object.fromEntries(bases) },
  }
}

/**
 * Read/Write fields edited here and not yet in the source: the document's
 * text differs from the base the link last agreed on.
 */
export function pendingWriteBacks(doc: PmNode, dvh: DvhDocsState, link: DvhLink): string[] {
  const texts = localFieldTexts(doc, dvh)
  return (dvh.model?.fields ?? [])
    .filter((f) => f.access === 'readwrite' && link.targets.includes(f.id))
    .filter((f) => {
      const base = linkFieldBase(link, f.id)
      return base !== undefined && texts.get(f.id) !== base.text
    })
    .map((f) => f.id)
}

/**
 * Pulls the linked workbook's field values into the model and every field
 * occurrence, and its collections into the model (tables are refreshed by the
 * caller, see dvh-tables updateLinkFromSource).
 */
export function updateFromSource(
  editor: Editor,
  dvh: DvhDocsState,
  link: DvhLink,
  source: DvhSource,
  options: { history?: boolean } = {},
): SourceUpdate {
  if (!dvh.model) return { fields: 0, collections: [], localEdits: [], conflicts: [] }
  const incoming = new Map(source.model.fields.map((f) => [f.id, f]))
  const changes: ChangeSet['changes'] = []
  const updatedCollections: string[] = []
  for (const next of source.model.collections) {
    if (!link.targets.includes(next.id)) continue
    const current = dvh.model.collections.find((c) => c.id === next.id)
    if (current && contentHash(current) === contentHash(next)) continue
    // the whole data, so the history can restore this collection to any point (P5)
    changes.push({
      objectId: next.id,
      path: 'data',
      before: current ? collectionData(current) : null,
      after: collectionData(next),
    })
    dvh.model.collections = [
      ...dvh.model.collections.filter((c) => c.id !== next.id),
      structuredClone(next),
    ]
    updatedCollections.push(next.id)
  }
  const texts = new Map<string, string>()
  const local = localFieldTexts(editor.state.doc, dvh)
  const bases = new Map<string, FieldBase>()
  const localEdits: string[] = []
  const conflicts: FieldConflict[] = []
  for (const field of dvh.model.fields) {
    const next = link.targets.includes(field.id) ? incoming.get(field.id) : undefined
    if (!next) continue
    const remote = fieldBaseOf(next)
    const base = linkFieldBase(link, field.id)
    // a read-only field is never edited here: the source always wins
    const state =
      field.access === 'read' ? 'source' : threeWay(base, local.get(field.id) ?? '', remote)
    if (state === 'local') {
      localEdits.push(field.id)
      continue
    }
    if (state === 'conflict') {
      conflicts.push({ fieldId: field.id, local: local.get(field.id) ?? '', source: remote, base })
      continue
    }
    bases.set(field.id, remote)
    texts.set(field.id, remote.text)
    if (next.value === field.value) continue
    changes.push({ objectId: field.id, path: 'value', before: field.value, after: next.value })
    field.value = next.value
  }
  link.lastSync = {
    revision: (link.lastSync?.revision ?? 0) + 1,
    hash: modelHash(source.model),
    at: new Date().toISOString(),
    fields: { ...(link.lastSync?.fields ?? {}), ...Object.fromEntries(bases) },
  }
  if (options.history === false) recordCoalesced(dvh, 'Link.Update', changes)
  else record(dvh, 'Link.Update', changes, 'link', true)
  const tr = editor.state.tr
  const changed = replaceFieldTexts(tr, texts)
  // automatic updates stay out of the user's undo history
  if (options.history === false) tr.setMeta('addToHistory', false)
  if (tr.docChanged) editor.view.dispatch(tr)
  return { fields: changed, collections: updatedCollections, localEdits, conflicts }
}

/**
 * Shows `text` in every occurrence of a field and makes it the model value
 * (a pulled source value, or a conflict settled for the source).
 */
export function applyFieldText(
  editor: Editor,
  dvh: DvhDocsState,
  fieldId: string,
  text: string,
  action: string,
  source: ChangeSet['source'],
): void {
  const field = dvh.model?.fields.find((f) => f.id === fieldId)
  if (!field) return
  const value = fieldValueFromText(field.type, text)
  const tr = editor.state.tr
  replaceFieldTexts(tr, new Map([[fieldId, text]]))
  if (tr.docChanged) editor.view.dispatch(tr)
  setFieldValue(dvh, field, value, action, source)
}

export type LinkUpdateMode = DvhLink['update']

/** Manual, on open, or automatic (watched file and live channel). */
export function setLinkUpdateMode(dvh: DvhDocsState, linkId: string, mode: LinkUpdateMode): void {
  const link = dvh.model?.links.find((l) => l.id === linkId)
  if (!link || link.update === mode) return
  const before = link.update
  link.update = mode
  record(
    dvh,
    'Link.SetMode',
    [{ objectId: link.id, path: 'update', before, after: mode }],
    'ui',
    true,
  )
}

/** Records a change set made by another DVH module (tables); the document becomes dirty. */
export function recordDvhDocsChange(
  dvh: DvhDocsState,
  action: string,
  changes: ChangeSet['changes'],
  source: ChangeSet['source'] = 'ui',
): void {
  record(dvh, action, changes, source, true)
}

/** The history to show: what the file holds plus this session's change sets, newest first. */
export function historyEntries(dvh: DvhDocsState | null): ChangeSet[] {
  return [...(dvh?.history?.changes ?? []), ...(dvh?.pending ?? [])].reverse()
}

/**
 * The model and history parts for saveDocx, or undefined when nothing DVH
 * changed. Not consuming: recovery copies build bytes too, and a completed save
 * reparses the document, which reloads this state from the saved parts.
 */
export function dvhDocsCustomXmlParts(
  parsed: ParsedDocFull,
  editor: Editor,
): CustomXmlPartWrite[] | undefined {
  const dvh = states.get(parsed)
  if (!dvh?.model) return undefined
  syncModelFromEditor(editor, dvh)
  // a compaction leaves no pending change but a new history to write
  if (dvh.pending.length === 0 && !historyRewritten.has(dvh)) return undefined
  dvh.modelStoreItemId ??= newStoreItemId()
  dvh.historyStoreItemId ??= newStoreItemId()
  const history = historyForSave(dvh.history, dvh.pending, dvh.model)
  return [
    { ns: MODEL_NS, xml: serializeModelXml(dvh.model), storeItemId: dvh.modelStoreItemId },
    { ns: HISTORY_NS, xml: serializeHistoryXml(history), storeItemId: dvh.historyStoreItemId },
  ]
}

/** The source model parsed from a dvhReadSource reply, or null. */
export function sourceFromRead(
  read: { path: string; modelXml: string | null } | null,
): DvhSource | null {
  if (!read?.modelXml) return null
  return { path: read.path, model: parseModelXml(read.modelXml) }
}

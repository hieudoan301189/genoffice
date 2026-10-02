/**
 * DVH links in Docs (P3), whether or not the Smart Data panel is open:
 *
 * - On open: links set to "on open" or "automatic" pull from their workbook.
 * - Automatic: the workbook file is watched (a save in DVH Sheets or Excel
 *   updates the document), and a workbook open in Sheets feeds the document
 *   live through the main process, without saving.
 * - Read/Write fields: a field edited here is written back to the workbook —
 *   to its Sheets tab when it is open, else into the file. Automatic links do
 *   it shortly after typing stops; other links on "Write to source".
 * - Conflicts: a field changed on both sides since the last sync is left as
 *   it is here and listed; the user keeps this side, keeps the source, or
 *   looks at the history first.
 * - Finding the source: relative path, last known path, then the project's
 *   files and the search index by docId. One match relinks silently; two or
 *   more are reported so the user picks.
 *
 * Automatic updates never overwrite a table edited by hand (the link shows
 * "edited" until a manual refresh) and stay out of the undo history.
 */
import type { Editor } from '@tiptap/core'
import {
  fieldValueFromText,
  modelHash,
  type DvhLink,
  type FieldBase,
  type WriteFieldsRequest,
  type WriteFieldsResult,
} from '@genoffice/dvh-model'
import {
  applyFieldText,
  linkFieldBase,
  linkIsStale,
  linkSourcePaths,
  localFieldTexts,
  pendingWriteBacks,
  recordDvhDocsChange,
  relocateLink,
  setLinkFieldBases,
  sourceFromRead,
  type DvhDocsState,
  type DvhSource,
  type FieldConflict,
} from './dvh-smart-data'
import { updateLinkFromSource, type LinkUpdate } from './dvh-tables'

export type LinkStatus = 'current' | 'stale' | 'missing' | 'edited' | 'conflict' | 'ambiguous'

type SourceRead = { path: string; modelXml: string | null } | null

export interface AutoLinkHost {
  editor(): Editor | null
  dvh(): DvhDocsState | null
  filePath(): string | null
  readSource(path: string): Promise<SourceRead>
  watch(paths: string[]): void
  /** write-back transport (main process); absent: Read/Write links stay read-only */
  writeFields?(request: WriteFieldsRequest): Promise<WriteFieldsResult>
  /** workbooks with this docId elsewhere on disk (project files, search index) */
  findSource?(docId: string): Promise<string[]>
}

const statuses = new Map<string, LinkStatus>()
const statusListeners = new Set<() => void>()
/** linkId → fieldId → conflict waiting for the user */
const conflicts = new Map<string, Map<string, FieldConflict>>()
/** linkId → workbooks that all carry the link's docId (the user picks one) */
const ambiguous = new Map<string, string[]>()

const emit = () => {
  for (const listener of statusListeners) listener()
}

export function linkStatus(linkId: string): LinkStatus | undefined {
  return statuses.get(linkId)
}

export function setLinkStatus(linkId: string, status: LinkStatus): void {
  if (statuses.get(linkId) === status) return
  statuses.set(linkId, status)
  emit()
}

export function onLinkStatus(listener: () => void): () => void {
  statusListeners.add(listener)
  return () => statusListeners.delete(listener)
}

/** The fields of a link changed on both sides, waiting for a decision. */
export function linkConflicts(linkId: string): FieldConflict[] {
  return [...(conflicts.get(linkId)?.values() ?? [])]
}

/** The workbooks a link could be reading when its docId is found more than once. */
export function ambiguousSources(linkId: string): string[] {
  return ambiguous.get(linkId) ?? []
}

function setConflicts(linkId: string, list: readonly FieldConflict[], replace: boolean): void {
  const map = replace ? new Map<string, FieldConflict>() : (conflicts.get(linkId) ?? new Map())
  for (const c of list) map.set(c.fieldId, c)
  if (map.size === 0) conflicts.delete(linkId)
  else conflicts.set(linkId, map)
}

function clearConflict(linkId: string, fieldId: string): void {
  const map = conflicts.get(linkId)
  map?.delete(fieldId)
  if (map?.size === 0) conflicts.delete(linkId)
}

/** The status a link shows after an update: conflicts first, then hand-edited tables. */
function settledStatus(linkId: string, editedTables: number): LinkStatus {
  if (conflicts.has(linkId)) return 'conflict'
  return editedTables > 0 ? 'edited' : 'current'
}

export type SourceLookup =
  | { kind: 'found'; source: DvhSource; relocated: boolean }
  | { kind: 'ambiguous'; paths: string[] }
  | { kind: 'missing' }

async function readMatching(
  link: DvhLink,
  path: string,
  read: (path: string) => Promise<SourceRead>,
): Promise<DvhSource | null> {
  try {
    const source = sourceFromRead(await read(path))
    return source && source.model.docId === link.source.docId ? source : null
  } catch {
    return null
  }
}

/**
 * Finds the linked workbook: relative path, last known path, then by docId
 * (project files, search index). A single docId match relinks the link to it.
 */
export async function resolveLinkedSource(
  dvh: DvhDocsState,
  link: DvhLink,
  docPath: string | null,
  read: (path: string) => Promise<SourceRead>,
  find?: (docId: string) => Promise<string[]>,
): Promise<SourceLookup> {
  const known = linkSourcePaths(link, docPath)
  for (const path of known) {
    const source = await readMatching(link, path, read)
    if (source) {
      ambiguous.delete(link.id)
      // found through relPath while the absolute path is stale (the folder moved)
      if (source.path !== link.source.path) relocateLink(dvh, link, source.path, docPath)
      return { kind: 'found', source, relocated: false }
    }
  }
  if (!find) return { kind: 'missing' }
  let candidates: string[]
  try {
    const lower = new Set(known.map((p) => p.toLowerCase()))
    candidates = (await find(link.source.docId)).filter((p) => !lower.has(p.toLowerCase()))
  } catch {
    return { kind: 'missing' }
  }
  const sources: DvhSource[] = []
  for (const path of candidates) {
    const source = await readMatching(link, path, read)
    if (source) sources.push(source)
  }
  if (sources.length > 1) {
    ambiguous.set(
      link.id,
      sources.map((s) => s.path),
    )
    return { kind: 'ambiguous', paths: sources.map((s) => s.path) }
  }
  ambiguous.delete(link.id)
  if (sources.length === 0) return { kind: 'missing' }
  relocateLink(dvh, link, sources[0]!.path, docPath)
  return { kind: 'found', source: sources[0]!, relocated: true }
}

/** The first candidate path that still holds the linked workbook (no search, no relink). */
export async function readLinkedSource(
  link: DvhLink,
  docPath: string | null,
  read: (path: string) => Promise<SourceRead>,
): Promise<DvhSource | null> {
  for (const path of linkSourcePaths(link, docPath)) {
    const source = await readMatching(link, path, read)
    if (source) return source
  }
  return null
}

/** Applies a workbook to a link and records the link's status and conflicts. */
export function applySource(
  editor: Editor,
  dvh: DvhDocsState,
  link: DvhLink,
  source: DvhSource,
  auto: boolean,
): LinkUpdate {
  const result = updateLinkFromSource(editor, dvh, link, source, { auto })
  // an update that ran re-judges every field; one skipped (source unchanged) keeps the list
  if (result.conflicts) setConflicts(link.id, result.conflicts, true)
  setLinkStatus(link.id, settledStatus(link.id, result.editedTables.length))
  return result
}

let host: AutoLinkHost | null = null
/** one update at a time: a live burst must not interleave with a file read or a write */
let queue: Promise<unknown> = Promise.resolve()
const enqueue = <T>(job: () => Promise<T> | T): Promise<T | undefined> => {
  const run = queue.then(job, job)
  queue = run.catch((error) => console.warn('DVH link:', error))
  return run.catch((error) => {
    console.warn('DVH link:', error)
    return undefined
  })
}

const autoLinks = (dvh: DvhDocsState | null) =>
  (dvh?.model?.links ?? []).filter((l) => l.update === 'auto')

/** Tells the main process which workbooks the automatic links read. */
export function refreshAutoWatch(): void {
  if (!host) return
  const docPath = host.filePath()
  host.watch([...new Set(autoLinks(host.dvh()).flatMap((l) => linkSourcePaths(l, docPath)))])
}

/** Reads every link's workbook: applies "on open" and "automatic" ones, flags stale manual ones. */
export function checkLinksOnOpen(): Promise<unknown> {
  return enqueue(async () => {
    const current = host
    const dvh = current?.dvh()
    const editor = current?.editor()
    if (!current || !dvh?.model || !editor) return
    let relocated = false
    for (const link of dvh.model.links) {
      const found = await resolveLinkedSource(
        dvh,
        link,
        current.filePath(),
        current.readSource,
        current.findSource,
      )
      if (found.kind !== 'found') {
        setLinkStatus(link.id, found.kind)
        continue
      }
      relocated ||= found.relocated
      if (link.update === 'manual') {
        setLinkStatus(link.id, linkIsStale(link, found.source) ? 'stale' : 'current')
      } else {
        applySource(editor, dvh, link, found.source, true)
      }
    }
    if (relocated) refreshAutoWatch()
  })
}

function onFileChanged(path: string): void {
  void enqueue(async () => {
    const current = host
    const dvh = current?.dvh()
    const editor = current?.editor()
    if (!current || !dvh || !editor) return
    const docPath = current.filePath()
    for (const link of autoLinks(dvh)) {
      if (!linkSourcePaths(link, docPath).some((p) => p.toLowerCase() === path.toLowerCase()))
        continue
      const source = sourceFromRead(await current.readSource(path))
      if (source && source.model.docId === link.source.docId)
        applySource(editor, dvh, link, source, true)
    }
  })
}

function onLive(payload: { docId: string; path: string | null; modelXml: string }): void {
  void enqueue(() => {
    const current = host
    const dvh = current?.dvh()
    const editor = current?.editor()
    if (!current || !dvh || !editor) return
    const links = autoLinks(dvh).filter((l) => l.source.docId === payload.docId)
    if (links.length === 0) return
    const source = sourceFromRead({ path: payload.path ?? '', modelXml: payload.modelXml })
    if (!source) return
    for (const link of links) applySource(editor, dvh, link, source, true)
  })
}

export interface WriteBackOutcome {
  readonly written: number
  readonly conflicts: number
  /** fields the source could not take: their cell holds a formula, or they are gone */
  readonly refused: readonly string[]
  /** nothing could be written: the workbook was not found or is locked */
  readonly error?: string
}

const fileName = (path: string | null) => (path ? (path.split(/[\\/]/).pop() ?? path) : undefined)

async function writeBackNow(
  current: AutoLinkHost,
  dvh: DvhDocsState,
  editor: Editor,
  link: DvhLink,
  fieldIds: readonly string[],
  force: boolean,
): Promise<WriteBackOutcome> {
  const model = dvh.model
  if (!model || !current.writeFields || fieldIds.length === 0) {
    return { written: 0, conflicts: 0, refused: [] }
  }
  const texts = localFieldTexts(editor.state.doc, dvh)
  const writes = fieldIds.flatMap((id) => {
    const field = model.fields.find((f) => f.id === id && f.access === 'readwrite')
    if (!field || !link.targets.includes(id)) return []
    const text = texts.get(id) ?? ''
    return [
      {
        fieldId: id,
        value: fieldValueFromText(field.type, text),
        expected: linkFieldBase(link, id) ?? null,
        force,
      },
    ]
  })
  if (writes.length === 0) return { written: 0, conflicts: 0, refused: [] }
  const docPath = current.filePath()
  const result = await current.writeFields({
    docId: link.source.docId,
    paths: linkSourcePaths(link, docPath),
    writes,
    origin: { docId: model.docId, ...(fileName(docPath) ? { name: fileName(docPath) } : {}) },
  })
  if (result.via === 'none') {
    setLinkStatus(link.id, result.path === null ? 'missing' : (statuses.get(link.id) ?? 'current'))
    return { written: 0, conflicts: 0, refused: [], error: result.error ?? 'not written' }
  }
  if (result.path && result.path !== link.source.path) relocateLink(dvh, link, result.path, docPath)
  const bases = new Map<string, FieldBase>()
  const changes: { objectId: string; path: string; before: unknown; after: unknown }[] = []
  const found: FieldConflict[] = []
  const refused: string[] = []
  for (const r of result.results) {
    const write = writes.find((w) => w.fieldId === r.fieldId)
    const field = model.fields.find((f) => f.id === r.fieldId)
    if (!write || !field) continue
    if (r.status === 'written' && r.current) {
      bases.set(r.fieldId, r.current)
      changes.push({
        objectId: r.fieldId,
        path: 'value',
        before: write.expected?.text ?? null,
        after: write.value,
      })
      field.value = write.value
      clearConflict(link.id, r.fieldId)
    } else if (r.status === 'conflict' && r.current) {
      found.push({
        fieldId: r.fieldId,
        local: texts.get(r.fieldId) ?? '',
        source: r.current,
        base: write.expected ?? undefined,
      })
    } else {
      refused.push(r.fieldId)
    }
  }
  setLinkFieldBases(link, bases)
  if (result.via === 'file' && result.path && bases.size > 0) {
    await advanceSyncHash(current, link, result.path, writes, bases)
  }
  // one change set per write-back, even when the model value already matched (synced on save)
  recordDvhDocsChange(dvh, 'Link.WriteBack', changes, 'link')
  setConflicts(link.id, found, false)
  setLinkStatus(link.id, settledStatus(link.id, 0))
  return { written: bases.size, conflicts: found.length, refused }
}

/**
 * After writing into the file, the workbook differs from what the link last
 * pulled only by our own writes: the link stays current (a manual link would
 * otherwise show "stale" for its own edit). Anything else changed in the
 * meantime keeps the old hash, so the link still reports it.
 */
async function advanceSyncHash(
  current: AutoLinkHost,
  link: DvhLink,
  path: string,
  writes: readonly { fieldId: string; expected: FieldBase | null }[],
  written: ReadonlyMap<string, FieldBase>,
): Promise<void> {
  let source: DvhSource | null
  try {
    source = sourceFromRead(await current.readSource(path))
  } catch {
    return
  }
  if (!source || !link.lastSync) return
  const before = structuredClone(source.model)
  for (const field of before.fields) {
    const expected = writes.find((w) => w.fieldId === field.id)?.expected
    if (!written.has(field.id) || !expected) continue
    field.value = fieldValueFromText(field.type, expected.text)
    // revision 0 is a field no DVH writer touched yet: no rev attribute
    if (expected.rev === 0) delete field.rev
    else field.rev = expected.rev
  }
  if (modelHash(before) === link.lastSync.hash) {
    link.lastSync = { ...link.lastSync, hash: modelHash(source.model) }
  }
}

/**
 * Writes a link's fields back to its workbook: the given fields, or every
 * field edited here since the last sync. `force` keeps this side over a
 * source that moved on (the "keep mine" answer to a conflict).
 */
export function writeBackLink(
  linkId: string,
  options: { fieldIds?: readonly string[]; force?: boolean } = {},
): Promise<WriteBackOutcome | undefined> {
  return enqueue(async () => {
    const current = host
    const dvh = current?.dvh()
    const editor = current?.editor()
    const link = dvh?.model?.links.find((l) => l.id === linkId)
    if (!current || !dvh || !editor || !link) return undefined
    const fieldIds = options.fieldIds ?? pendingWriteBacks(editor.state.doc, dvh, link)
    return writeBackNow(current, dvh, editor, link, fieldIds, options.force === true)
  })
}

/** Fields edited here, not yet written to the link's workbook. */
export function linkPendingWrites(editor: Editor, dvh: DvhDocsState, link: DvhLink): string[] {
  return pendingWriteBacks(editor.state.doc, dvh, link).filter(
    (id) => !conflicts.get(link.id)?.has(id),
  )
}

/** Settles a conflict: keep this document's text (written over the source) or take the source's. */
export function resolveConflict(
  linkId: string,
  fieldId: string,
  keep: 'mine' | 'source',
): Promise<WriteBackOutcome | undefined> {
  if (keep === 'mine') return writeBackLink(linkId, { fieldIds: [fieldId], force: true })
  return enqueue(() => {
    const current = host
    const dvh = current?.dvh()
    const editor = current?.editor()
    const link = dvh?.model?.links.find((l) => l.id === linkId)
    const conflict = conflicts.get(linkId)?.get(fieldId)
    if (!dvh || !editor || !link || !conflict) return undefined
    applyFieldText(editor, dvh, fieldId, conflict.source.text, 'Link.ResolveConflict', 'link')
    setLinkFieldBases(link, new Map([[fieldId, conflict.source]]))
    clearConflict(linkId, fieldId)
    setLinkStatus(linkId, settledStatus(linkId, 0))
    return { written: 0, conflicts: 0, refused: [] }
  })
}

/** after typing stops, automatic links write their edited fields back */
const AUTO_WRITE_DELAY_MS = 1200
let autoWriteTimer: ReturnType<typeof setTimeout> | undefined

/** Called on document edits: schedules the write-back of automatic links. */
export function noteDocumentEdited(): void {
  clearTimeout(autoWriteTimer)
  autoWriteTimer = setTimeout(() => {
    autoWriteTimer = undefined
    const current = host
    const dvh = current?.dvh()
    const editor = current?.editor()
    if (!current?.writeFields || !dvh || !editor) return
    for (const link of autoLinks(dvh)) {
      if (linkPendingWrites(editor, dvh, link).length === 0) continue
      void writeBackLink(link.id, { fieldIds: linkPendingWrites(editor, dvh, link) })
    }
  }, AUTO_WRITE_DELAY_MS)
}

/** Starts the link runner for the open document; returns the stop function. */
export function installAutoLinks(next: AutoLinkHost): () => void {
  host = next
  const offFile = window.desktop.onDvhSourceChanged?.(onFileChanged)
  const offLive = window.desktop.onDvhLive?.(onLive)
  return () => {
    if (host === next) host = null
    clearTimeout(autoWriteTimer)
    offFile?.()
    offLive?.()
    next.watch([])
  }
}

/** Test hook: forget conflicts, choices and statuses between documents. */
export function resetLinkState(): void {
  statuses.clear()
  conflicts.clear()
  ambiguous.clear()
}

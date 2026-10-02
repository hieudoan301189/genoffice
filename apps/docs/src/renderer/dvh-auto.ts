/**
 * DVH automatic links in Docs (P3). Runs whether or not the Smart Data panel
 * is open:
 *
 * - On open: links set to "on open" or "automatic" pull from their workbook.
 * - Automatic: the workbook file is watched (a save in DVH Sheets or Excel
 *   updates the document), and a workbook open in Sheets feeds the document
 *   live through the main process, without saving.
 *
 * Automatic updates never overwrite a table edited by hand (the link shows
 * "edited" until a manual refresh) and stay out of the undo history.
 */
import type { Editor } from '@tiptap/core'
import type { DvhLink } from '@genoffice/dvh-model'
import {
  linkIsStale,
  linkSourcePaths,
  sourceFromRead,
  type DvhDocsState,
  type DvhSource,
} from './dvh-smart-data'
import { updateLinkFromSource, type LinkUpdate } from './dvh-tables'

export type LinkStatus = 'current' | 'stale' | 'missing' | 'edited'

type SourceRead = { path: string; modelXml: string | null } | null

export interface AutoLinkHost {
  editor(): Editor | null
  dvh(): DvhDocsState | null
  filePath(): string | null
  readSource(path: string): Promise<SourceRead>
  watch(paths: string[]): void
}

const statuses = new Map<string, LinkStatus>()
const statusListeners = new Set<() => void>()

export function linkStatus(linkId: string): LinkStatus | undefined {
  return statuses.get(linkId)
}

export function setLinkStatus(linkId: string, status: LinkStatus): void {
  if (statuses.get(linkId) === status) return
  statuses.set(linkId, status)
  for (const listener of statusListeners) listener()
}

export function onLinkStatus(listener: () => void): () => void {
  statusListeners.add(listener)
  return () => statusListeners.delete(listener)
}

/** The first candidate path that still holds the linked workbook. */
export async function readLinkedSource(
  link: DvhLink,
  docPath: string | null,
  read: (path: string) => Promise<SourceRead>,
): Promise<DvhSource | null> {
  for (const path of linkSourcePaths(link, docPath)) {
    try {
      const source = sourceFromRead(await read(path))
      if (source && source.model.docId === link.source.docId) return source
    } catch {
      // moved or unreadable: try the next candidate
    }
  }
  return null
}

/** Applies a workbook to a link and records the link's status. */
export function applySource(
  editor: Editor,
  dvh: DvhDocsState,
  link: DvhLink,
  source: DvhSource,
  auto: boolean,
): LinkUpdate {
  const result = updateLinkFromSource(editor, dvh, link, source, { auto })
  setLinkStatus(link.id, result.editedTables.length > 0 ? 'edited' : 'current')
  return result
}

let host: AutoLinkHost | null = null
/** one update at a time: a live burst must not interleave with a file read */
let queue: Promise<unknown> = Promise.resolve()
const enqueue = (job: () => Promise<unknown> | unknown) => {
  queue = queue.then(job, job).catch((error) => console.warn('DVH automatic link:', error))
  return queue
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
    for (const link of dvh.model.links) {
      const source = await readLinkedSource(link, current.filePath(), current.readSource)
      if (!source) {
        setLinkStatus(link.id, 'missing')
        continue
      }
      if (link.update === 'manual') {
        setLinkStatus(link.id, linkIsStale(link, source) ? 'stale' : 'current')
      } else {
        applySource(editor, dvh, link, source, true)
      }
    }
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

/** Starts automatic links for the open document; returns the stop function. */
export function installAutoLinks(next: AutoLinkHost): () => void {
  host = next
  const offFile = window.desktop.onDvhSourceChanged?.(onFileChanged)
  const offLive = window.desktop.onDvhLive?.(onLive)
  return () => {
    if (host === next) host = null
    offFile?.()
    offLive?.()
    next.watch([])
  }
}

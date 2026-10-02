// DVH automatic links (P3), main-process side.
//
// - A document lists the workbooks its automatic links read; each is watched
//   on disk (its folder, so Excel's and DVH Sheets' save-by-rename is seen) and
//   the document hears `docs:dvh-source-changed` once the file settles.
// - A workbook open in Sheets publishes its Smart Data after each edit; this
//   relays it to every subscribed document (`docs:dvh-live`), no save needed.
// - Field write-back (Read/Write links): a document's write goes to the Sheets
//   tab holding the workbook (`Data.SetField` there, undoable, saved with the
//   workbook) or, when no tab holds it, into the xlsx on disk.
import { randomUUID } from 'node:crypto'
import { watch, type FSWatcher } from 'node:fs'
import { readFile, stat } from 'node:fs/promises'
import { basename, dirname, extname } from 'node:path'
import { ipcMain, webContents, type WebContents } from 'electron'
import {
  DVH_LIVE_DOCS_CHANNEL,
  DVH_LIVE_PUBLISH_CHANNEL,
  DVH_SHEETS_SET_FIELDS_CHANNEL,
  DVH_SHEETS_SET_FIELDS_RESULT_CHANNEL,
  DVH_WRITE_FIELDS_CHANNEL,
  dvhLivePayloadSchema,
  sheetsSetFieldsReplySchema,
  writeFieldsRequestSchema,
  type SheetsSetFieldsReply,
  type WriteFieldsRequest,
  type WriteFieldsResult,
} from '@genoffice/dvh-model'
import { atomicWriteFile } from './atomic-write'
import { writeFieldsToXlsx } from './dvh-xlsx-write'

const SOURCE_CHANGED_CHANNEL = 'docs:dvh-source-changed'
/** writers save in steps (temp file, rename, metadata): report once the file is quiet */
const SETTLE_MS = 700

interface FolderWatch {
  watcher: FSWatcher
  /** lower-cased file name → full path */
  files: Map<string, string>
  timers: Map<string, ReturnType<typeof setTimeout>>
}

/** webContents id → the workbook paths its document watches */
const subscriptions = new Map<number, Set<string>>()
const folders = new Map<string, FolderWatch>()
const lastSeen = new Map<string, string>()
/** workbook docId → the Sheets webContents that last published it */
const publishers = new Map<string, number>()
/** a Sheets tab answers a write within this time, or the write goes to the file */
const LIVE_WRITE_TIMEOUT_MS = 5000
/** requestId → the tab asked and how to finish; only that tab's answer counts */
const pendingWrites = new Map<
  string,
  { sender: number; finish: (reply: SheetsSetFieldsReply | null) => void }
>()

const key = (path: string) => path.toLowerCase()

async function fingerprint(path: string): Promise<string | null> {
  try {
    const info = await stat(path)
    return `${info.size}:${info.mtimeMs}`
  } catch {
    return null
  }
}

/** Project data (P10) is not a file: its changes are announced by the project service. */
const isProjectUri = (path: string) => path.startsWith('dvh-project://')

/** Tells the documents watching a source that it changed (project data, P10). */
export function notifyDvhSourceChanged(path: string): void {
  notify(path)
}

let projectFieldWriter: ((request: WriteFieldsRequest) => WriteFieldsResult | null) | null = null

/** Field write-back for `dvh-project://` sources (P10). */
export function setProjectFieldWriter(
  writer: ((request: WriteFieldsRequest) => WriteFieldsResult | null) | null,
): void {
  projectFieldWriter = writer
}

function notify(path: string): void {
  for (const [id, paths] of subscriptions) {
    if (![...paths].some((p) => key(p) === key(path))) continue
    const target = webContents.fromId(id)
    if (target && !target.isDestroyed()) target.send(SOURCE_CHANGED_CHANNEL, path)
  }
}

function onFolderEvent(folder: FolderWatch, fileName: string | null): void {
  if (!fileName) return
  const path = folder.files.get(fileName.toLowerCase())
  if (!path) return
  clearTimeout(folder.timers.get(path))
  folder.timers.set(
    path,
    setTimeout(() => {
      folder.timers.delete(path)
      void fingerprint(path).then((print) => {
        // vanished mid-save, or a touch that did not change the bytes
        if (print === null || print === lastSeen.get(key(path))) return
        lastSeen.set(key(path), print)
        notify(path)
      })
    }, SETTLE_MS),
  )
}

/** Watches exactly the folders the current subscriptions need. */
function rewatch(): void {
  const wanted = new Map<string, Map<string, string>>()
  for (const paths of subscriptions.values()) {
    for (const path of paths) {
      if (isProjectUri(path)) continue
      const dir = dirname(path)
      const files = wanted.get(key(dir)) ?? new Map<string, string>()
      files.set(basename(path).toLowerCase(), path)
      wanted.set(key(dir), files)
    }
  }
  for (const [dir, folder] of folders) {
    if (wanted.has(dir)) continue
    folder.watcher.close()
    for (const timer of folder.timers.values()) clearTimeout(timer)
    folders.delete(dir)
  }
  for (const [dir, files] of wanted) {
    const existing = folders.get(dir)
    if (existing) {
      existing.files = files
      continue
    }
    const first = [...files.values()][0]!
    try {
      const folder: FolderWatch = {
        watcher: watch(dirname(first), { persistent: false }),
        files,
        timers: new Map(),
      }
      folder.watcher.on('change', (_event, fileName) =>
        onFolderEvent(folder, typeof fileName === 'string' ? fileName : null),
      )
      folder.watcher.on('error', () => {
        folder.watcher.close()
        folders.delete(dir)
      })
      folders.set(dir, folder)
      for (const path of files.values()) {
        void fingerprint(path).then((print) => {
          if (print) lastSeen.set(key(path), print)
        })
      }
    } catch {
      // folder gone or not watchable: the link stays manual until the document reopens
    }
  }
}

function subscribe(sender: WebContents, paths: readonly string[]): void {
  const id = sender.id
  if (!subscriptions.has(id)) {
    sender.once('destroyed', () => {
      subscriptions.delete(id)
      rewatch()
    })
  }
  subscriptions.set(id, new Set(paths))
  rewatch()
}

export function registerDvhLinkIpc(): void {
  // a document's automatic links: the workbooks to watch (empty list = none)
  ipcMain.handle('docs:dvh-watch', (event, paths: unknown) => {
    const list = Array.isArray(paths)
      ? paths
          .filter(
            (p): p is string => typeof p === 'string' && (/\.xls[xm]$/i.test(p) || isProjectUri(p)),
          )
          .slice(0, 64)
      : []
    subscribe(event.sender, list)
  })

  // Sheets → every subscribed document; the payload is validated before it is passed on
  ipcMain.on(DVH_LIVE_PUBLISH_CHANNEL, (event, payload: unknown) => {
    const parsed = dvhLivePayloadSchema.safeParse(payload)
    if (!parsed.success) return
    notePublisher(parsed.data.docId, event.sender)
    for (const id of subscriptions.keys()) {
      if (id === event.sender.id) continue
      const target = webContents.fromId(id)
      if (target && !target.isDestroyed()) target.send(DVH_LIVE_DOCS_CHANNEL, parsed.data)
    }
  })

  ipcMain.on(DVH_SHEETS_SET_FIELDS_RESULT_CHANNEL, (event, reply: unknown) => {
    const parsed = sheetsSetFieldsReplySchema.safeParse(reply)
    if (!parsed.success) return
    const pending = pendingWrites.get(parsed.data.requestId)
    if (pending && pending.sender === event.sender.id) pending.finish(parsed.data)
  })

  ipcMain.handle(DVH_WRITE_FIELDS_CHANNEL, async (_event, request: unknown) => {
    const parsed = writeFieldsRequestSchema.safeParse(request)
    if (!parsed.success) throw new Error('Invalid DVH write request.')
    return writeDvhFields(parsed.data)
  })
}

function notePublisher(docId: string, sender: WebContents): void {
  if (publishers.get(docId) === sender.id) return
  publishers.set(docId, sender.id)
  sender.once('destroyed', () => {
    if (publishers.get(docId) === sender.id) publishers.delete(docId)
  })
}

/** Sends the write to the Sheets tab that published the workbook; null when none answers for it. */
function writeLive(request: WriteFieldsRequest): Promise<SheetsSetFieldsReply | null> {
  const id = publishers.get(request.docId)
  const target = id === undefined ? null : webContents.fromId(id)
  if (!target || target.isDestroyed()) return Promise.resolve(null)
  const requestId = randomUUID()
  return new Promise((resolve) => {
    const timer = setTimeout(() => finish(null), LIVE_WRITE_TIMEOUT_MS)
    const finish = (reply: SheetsSetFieldsReply | null) => {
      clearTimeout(timer)
      pendingWrites.delete(requestId)
      resolve(reply && reply.handled ? reply : null)
    }
    pendingWrites.set(requestId, { sender: target.id, finish })
    target.send(DVH_SHEETS_SET_FIELDS_CHANNEL, {
      requestId,
      docId: request.docId,
      writes: request.writes,
      origin: request.origin,
    })
  })
}

/** Writes fields into the first candidate workbook on disk that is the linked one. */
async function writeFile(request: WriteFieldsRequest): Promise<WriteFieldsResult> {
  let lastError = 'The source workbook was not found.'
  for (const path of request.paths) {
    if (!/^\.xls[xm]$/i.test(extname(path))) continue
    let bytes: Uint8Array
    try {
      bytes = new Uint8Array(await readFile(path))
    } catch {
      continue
    }
    let out
    try {
      out = await writeFieldsToXlsx(bytes, request)
    } catch (error) {
      // another workbook at that path (moved, replaced): try the next candidate
      lastError = error instanceof Error ? error.message : String(error)
      continue
    }
    if (out.bytes) {
      try {
        await atomicWriteFile(path, Buffer.from(out.bytes))
      } catch (error) {
        // Excel holds the file open (Windows locks it): nothing was written
        return {
          via: 'none',
          path,
          results: [],
          error: error instanceof Error ? error.message : String(error),
        }
      }
    }
    return { via: 'file', path, results: out.results }
  }
  return { via: 'none', path: null, results: [], error: lastError }
}

export async function writeDvhFields(request: WriteFieldsRequest): Promise<WriteFieldsResult> {
  const project = projectFieldWriter?.(request)
  if (project) return project
  const live = await writeLive(request)
  if (live) return { via: 'live', path: live.path, results: live.results }
  return writeFile(request)
}

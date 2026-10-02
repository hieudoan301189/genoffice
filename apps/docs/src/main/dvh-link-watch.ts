// DVH automatic links (P3), main-process side.
//
// - A document lists the workbooks its automatic links read; each is watched
//   on disk (its folder, so Excel's and DVH Sheets' save-by-rename is seen) and
//   the document hears `docs:dvh-source-changed` once the file settles.
// - A workbook open in Sheets publishes its Smart Data after each edit; this
//   relays it to every subscribed document (`docs:dvh-live`), no save needed.
import { watch, type FSWatcher } from 'node:fs'
import { stat } from 'node:fs/promises'
import { basename, dirname } from 'node:path'
import { ipcMain, webContents, type WebContents } from 'electron'
import {
  DVH_LIVE_DOCS_CHANNEL,
  DVH_LIVE_PUBLISH_CHANNEL,
  dvhLivePayloadSchema,
} from '@genoffice/dvh-model'

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

const key = (path: string) => path.toLowerCase()

async function fingerprint(path: string): Promise<string | null> {
  try {
    const info = await stat(path)
    return `${info.size}:${info.mtimeMs}`
  } catch {
    return null
  }
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
          .filter((p): p is string => typeof p === 'string' && /\.xls[xm]$/i.test(p))
          .slice(0, 64)
      : []
    subscribe(event.sender, list)
  })

  // Sheets → every subscribed document; the payload is validated before it is passed on
  ipcMain.on(DVH_LIVE_PUBLISH_CHANNEL, (event, payload: unknown) => {
    const parsed = dvhLivePayloadSchema.safeParse(payload)
    if (!parsed.success) return
    for (const id of subscriptions.keys()) {
      if (id === event.sender.id) continue
      const target = webContents.fromId(id)
      if (target && !target.isDestroyed()) target.send(DVH_LIVE_DOCS_CHANNEL, parsed.data)
    }
  })
}

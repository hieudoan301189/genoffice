// Finding a linked workbook that moved (P3 source resolution). A link keeps
// the workbook's DVH docId; when neither its relative nor its last known path
// holds that workbook any more, the candidates are the files the project store
// knows (fileMap, kept up to date on renames) and the shell's file index
// (docId column), and each is confirmed by reading its model part.
import { readFile, stat } from 'node:fs/promises'
import { extname } from 'node:path'
import JSZip from 'jszip'
import { MODEL_NS } from '@genoffice/dvh-model'

const SOURCE_EXTENSIONS = new Set(['.xlsx', '.xlsm'])
const PEEK_EXTENSIONS = new Set(['.xlsx', '.xlsm', '.docx'])
/** files larger than this are not opened just to read their docId */
const MAX_PEEK_BYTES = 128 * 1024 * 1024
/** bound the work of one search: project files beyond this are not peeked */
const MAX_CANDIDATES = 2000

/** The DVH docId in a package's model part, or null (no part, not a zip). */
export async function peekDvhDocIdFromBytes(bytes: Uint8Array): Promise<string | null> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch {
    return null
  }
  for (const path of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(path)) continue
    const xml = await zip.file(path)!.async('string')
    const root = /<(?:[\w-]+:)?model\b[^>]*>/.exec(xml)
    if (!root || !root[0].includes(`"${MODEL_NS}"`)) continue
    return /\sdocId=["']([^"']+)["']/.exec(root[0])?.[1] ?? null
  }
  return null
}

const peeked = new Map<string, { print: string; docId: string | null }>()

/** The docId of a workbook or document on disk (cached by size + mtime). */
export async function peekDvhDocId(path: string): Promise<string | null> {
  if (!PEEK_EXTENSIONS.has(extname(path).toLowerCase())) return null
  try {
    const info = await stat(path)
    if (!info.isFile() || info.size > MAX_PEEK_BYTES) return null
    const print = `${info.size}:${info.mtimeMs}`
    const known = peeked.get(path)
    if (known?.print === print) return known.docId
    const docId = await peekDvhDocIdFromBytes(new Uint8Array(await readFile(path)))
    peeked.set(path, { print, docId })
    return docId
  } catch {
    return null
  }
}

/**
 * The workbooks among `candidates` whose docId is `docId`, in candidate order
 * (project files first, then the file index), without duplicates.
 */
export async function findDvhSources(
  docId: string,
  candidates: readonly string[],
): Promise<string[]> {
  const seen = new Set<string>()
  const found: string[] = []
  let checked = 0
  for (const path of candidates) {
    const key = path.toLowerCase()
    if (seen.has(key) || !SOURCE_EXTENSIONS.has(extname(path).toLowerCase())) continue
    seen.add(key)
    if (++checked > MAX_CANDIDATES) break
    if ((await peekDvhDocId(path)) === docId) found.push(path)
  }
  return found
}

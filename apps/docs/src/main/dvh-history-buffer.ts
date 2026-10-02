// In-session history buffer (P5): change sets made since the last save are
// mirrored to userData/dvh-history/<docId>.jsonl, so a crash or a forced quit
// does not lose them silently. A save clears the buffer (the change sets are
// in the file then); the next open of the same document finds what is left
// and offers it back.
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { changeSetSchema, type ChangeSet } from '@genoffice/dvh-model'

const DOC_ID = /^doc_[a-z2-7]{16}$/
/** a buffer larger than this is not read back (the file's own history is the record) */
const MAX_BUFFER_BYTES = 8 * 1024 * 1024

const bufferPath = (dir: string, docId: string) => join(dir, `${docId}.jsonl`)

/** Replaces the buffer of one document with the session's change sets (empty list: removes it). */
export async function writeHistoryBuffer(
  dir: string,
  docId: string,
  changes: readonly unknown[],
): Promise<void> {
  if (!DOC_ID.test(docId)) return
  const valid = changes.filter((c) => changeSetSchema.safeParse(c).success)
  if (valid.length === 0) {
    await rm(bufferPath(dir, docId), { force: true })
    return
  }
  await mkdir(dir, { recursive: true })
  const text = valid.map((c) => JSON.stringify(c)).join('\n')
  if (text.length > MAX_BUFFER_BYTES) return
  await writeFile(bufferPath(dir, docId), text, 'utf8')
}

/** The change sets left in a document's buffer (none when it was saved cleanly). */
export async function readHistoryBuffer(dir: string, docId: string): Promise<ChangeSet[]> {
  if (!DOC_ID.test(docId)) return []
  let text: string
  try {
    text = await readFile(bufferPath(dir, docId), 'utf8')
  } catch {
    return []
  }
  if (text.length > MAX_BUFFER_BYTES) return []
  const out: ChangeSet[] = []
  for (const line of text.split('\n')) {
    if (!line.trim()) continue
    try {
      const parsed = changeSetSchema.safeParse(JSON.parse(line))
      if (parsed.success) out.push(parsed.data)
    } catch {
      // a torn last line (crash mid-write) is skipped
    }
  }
  return out
}

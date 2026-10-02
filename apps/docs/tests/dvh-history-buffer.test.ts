/**
 * P5 crash safety: the session's change sets are mirrored to disk until the
 * next save; what a crashed session left is read back, a torn line skipped.
 */
import { appendFileSync, mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { readHistoryBuffer, writeHistoryBuffer } from '../src/main/dvh-history-buffer'

let dir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'dvh-buffer-'))
})
afterEach(() => rmSync(dir, { recursive: true, force: true }))

const DOC = 'doc_abcdefghijklmnop'
const cs = (id: string) => ({
  id,
  txId: `tx_${id}`,
  docId: DOC,
  at: '2026-10-02T00:00:00.000Z',
  source: 'ui',
  action: 'Data.SetField',
  changes: [{ objectId: 'f_a', path: 'value', before: 'A', after: 'B' }],
})

describe('history buffer', () => {
  it('round-trips valid change sets, skips torn lines, clears on an empty list', async () => {
    await writeHistoryBuffer(dir, DOC, [cs('cs_1'), { nonsense: true }, cs('cs_2')])
    appendFileSync(join(dir, `${DOC}.jsonl`), '\n{"id":"cs_3","txId"')
    expect((await readHistoryBuffer(dir, DOC)).map((c) => c.id)).toEqual(['cs_1', 'cs_2'])
    await writeHistoryBuffer(dir, DOC, [])
    expect(await readHistoryBuffer(dir, DOC)).toEqual([])
  })

  it('refuses ids that are not DVH document ids (no path tricks)', async () => {
    await writeHistoryBuffer(dir, '../../evil', [cs('cs_1')])
    expect(await readHistoryBuffer(dir, '../../evil')).toEqual([])
  })
})

/**
 * P3 source resolution: candidate files are confirmed by the docId in their
 * model part; other workbooks, plain files and duplicates are skipped.
 */
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { emptyModel, serializeModelXml } from '@genoffice/dvh-model'
import { findDvhSources, peekDvhDocId } from '../src/main/dvh-find-source'

let dir: string
beforeEach(() => {
  dir = mkdtempSync(join(tmpdir(), 'dvh-find-'))
})
afterEach(() => rmSync(dir, { recursive: true, force: true }))

async function workbook(name: string, docId: string | null): Promise<string> {
  const zip = new JSZip()
  zip.file('xl/workbook.xml', '<workbook/>')
  // Excel renumbers parts: the model is found by namespace, not by item number
  zip.file('customXml/item1.xml', '<?xml version="1.0"?><other xmlns="urn:x"/>')
  if (docId) zip.file('customXml/item3.xml', serializeModelXml(emptyModel(docId)))
  const path = join(dir, name)
  writeFileSync(path, await zip.generateAsync({ type: 'uint8array' }))
  return path
}

describe('findDvhSources', () => {
  it('keeps only workbooks carrying the docId, once each', async () => {
    const a = await workbook('a.xlsx', 'doc_target')
    const b = await workbook('Moved Copy.xlsm', 'doc_target')
    const other = await workbook('other.xlsx', 'doc_other')
    const plain = await workbook('plain.xlsx', null)
    const text = join(dir, 'notes.txt')
    writeFileSync(text, 'doc_target')
    expect(await peekDvhDocId(a)).toBe('doc_target')
    expect(await peekDvhDocId(plain)).toBeNull()
    expect(await peekDvhDocId(join(dir, 'missing.xlsx'))).toBeNull()
    expect(
      await findDvhSources('doc_target', [other, a, plain, text, a, b, join(dir, 'gone.xlsx')]),
    ).toEqual([a, b])
  })
})

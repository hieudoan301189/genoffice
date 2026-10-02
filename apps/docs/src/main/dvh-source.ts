// DVH Smart Data link source (P1, ADR D8 minimal form): a document reads the
// model part of a linked workbook straight from disk, without opening it in
// Sheets. Sheets syncs field values from their bound cells on save, so the
// model part already carries current values.
import { readFile, stat } from 'node:fs/promises'
import { extname } from 'node:path'
import { readCustomXmlPart } from '@genoffice/docx-engine'
import { MODEL_NS } from '@genoffice/dvh-model'

const SOURCE_EXTENSIONS = new Set(['.xlsx', '.xlsm'])
/** a workbook larger than this is not read into memory just for its model part */
const MAX_SOURCE_BYTES = 256 * 1024 * 1024

export interface DvhSourceRead {
  readonly path: string
  /** the workbook's model part, null when it has no DVH Smart Data */
  readonly modelXml: string | null
}

export async function readDvhSource(path: string): Promise<DvhSourceRead> {
  if (!SOURCE_EXTENSIONS.has(extname(path).toLowerCase())) {
    throw new Error('A Smart Data source must be an .xlsx workbook.')
  }
  const info = await stat(path)
  if (!info.isFile() || info.size > MAX_SOURCE_BYTES) throw new Error('Workbook not readable.')
  const part = await readCustomXmlPart(new Uint8Array(await readFile(path)), MODEL_NS)
  return { path, modelXml: part?.xml ?? null }
}

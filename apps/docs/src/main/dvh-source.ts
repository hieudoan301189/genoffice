// DVH Smart Data link source (P1, ADR D8 minimal form): a document reads the
// model part of a linked workbook straight from disk, without opening it in
// Sheets. Values come from the cells the hidden DVH names point at, so a
// workbook last saved by Microsoft Excel (stale model part) still reads right.
import { readFile, stat } from 'node:fs/promises'
import { extname } from 'node:path'
import { readCustomXmlPart } from '@genoffice/docx-engine'
import { MODEL_NS, parseModelXml, serializeModelXml } from '@genoffice/dvh-model'
import { modelFromCells } from './dvh-xlsx-cells'

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
  const bytes = new Uint8Array(await readFile(path))
  const part = await readCustomXmlPart(bytes, MODEL_NS)
  if (!part) return { path, modelXml: null }
  // the cells are the truth: a workbook saved by Excel keeps a stale model part
  try {
    const model = await modelFromCells(bytes, parseModelXml(part.xml))
    return { path, modelXml: serializeModelXml(model) }
  } catch {
    return { path, modelXml: part.xml }
  }
}

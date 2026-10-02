// Smart Template jobs in the main process (P6): batch generation to a folder
// or a zip with its index, PDF through the headless export path, image
// fields loaded from disk, previews of one record, DVH-Tool conversion, and
// the QLCL workbook exchange. The engine itself is @genoffice/dvh-template.
import { spawn } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { dirname, extname, isAbsolute, join } from 'node:path'
import { atomicWriteFile } from './atomic-write'
import {
  batchIndexCsv,
  generateBatch,
  packageBatch,
  planBatch,
  type BatchSpec,
  type GeneratedDocument,
  type TemplateImage,
} from '@genoffice/dvh-template'
import { fieldText, type Scalar } from '@genoffice/dvh-model'

const IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'bmp'])
const MAX_IMAGE_BYTES = 20 * 1024 * 1024

/**
 * Image field values: a data URL, or a picture path (absolute, or relative to
 * the template's folder). Anything else is reported by the engine.
 */
export function imageLoader(baseDir: string | null): (value: Scalar) => TemplateImage | null {
  const cache = new Map<string, TemplateImage | null>()
  return (value) => {
    const text = fieldText(value).trim()
    if (!text) return null
    if (cache.has(text)) return cache.get(text)!
    let image: TemplateImage | null = null
    const data = /^data:image\/(png|jpe?g|gif|bmp);base64,([A-Za-z0-9+/=]+)$/.exec(text)
    if (data) {
      image = {
        ext: data[1] === 'jpeg' ? 'jpg' : data[1]!,
        bytes: new Uint8Array(Buffer.from(data[2]!, 'base64')),
      }
    } else {
      const path = isAbsolute(text) ? text : baseDir ? join(baseDir, text) : null
      const ext = path ? extname(path).slice(1).toLowerCase() : ''
      if (path && IMAGE_EXTS.has(ext) && existsSync(path)) {
        try {
          // synchronous read keeps the engine's loader simple; pictures are small
          const bytes = readFileSync(path)
          if (bytes.length <= MAX_IMAGE_BYTES)
            image = { ext: ext === 'jpeg' ? 'jpg' : ext, bytes: new Uint8Array(bytes) }
        } catch {
          image = null
        }
      }
    }
    cache.set(text, image)
    return image
  }
}

/**
 * Converts one docx to PDF through the app's own headless export
 * (docs/headless-pdf-export.md): the same renderer as File → Export PDF.
 */
export function exportPdfHeadless(
  input: string,
  output: string,
  launcher: { execPath: string; appArgs: readonly string[] },
): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(
      launcher.execPath,
      [...launcher.appArgs, '--headless-export', input, '--to', 'pdf', '--out', output, '--json'],
      { stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env, ELECTRON_RUN_AS_NODE: '' } },
    )
    let out = ''
    child.stdout.on('data', (chunk: Buffer) => (out += chunk.toString('utf8')))
    child.on('error', reject)
    child.on('close', (code) => {
      if (code === 0) resolve()
      else reject(new Error(out.trim() || `PDF export failed (${code})`))
    })
  })
}

export type BatchOutput =
  | { readonly kind: 'folder'; readonly dir: string }
  | { readonly kind: 'zip'; readonly path: string }

export interface BatchRunResult {
  readonly count: number
  readonly output: string
  readonly warnings: readonly string[]
}

/** Names that would overwrite existing files in a folder output. */
export function existingNames(
  dir: string,
  names: readonly string[],
  format: 'docx' | 'pdf',
): string[] {
  return names
    .map((n) => (format === 'pdf' ? n.replace(/\.docx$/i, '.pdf') : n))
    .filter((n) => existsSync(join(dir, n)))
}

/** Generates, converts (PDF) and writes a batch; documents are written as they are made. */
export async function runBatch(
  spec: BatchSpec,
  output: BatchOutput,
  format: 'docx' | 'pdf',
  launcher: { execPath: string; appArgs: readonly string[] },
  tempDir: string,
): Promise<BatchRunResult> {
  const plan = planBatch(spec)
  const documents: GeneratedDocument[] = await generateBatch(spec, plan)
  const warnings = [
    ...plan.warnings,
    ...documents.flatMap((d) => d.warnings.map((w) => `${d.name}: ${w}`)),
  ]
  let files: { name: string; bytes: Uint8Array; record?: GeneratedDocument['record'] }[] = documents
  if (format === 'pdf') {
    const work = join(tempDir, `dvh-batch-${Date.now()}`)
    await mkdir(work, { recursive: true })
    try {
      files = []
      for (const doc of documents) {
        const docx = join(work, doc.name)
        const pdf = docx.replace(/\.docx$/i, '.pdf')
        await writeFile(docx, doc.bytes)
        await exportPdfHeadless(docx, pdf, launcher)
        files.push({
          name: doc.name.replace(/\.docx$/i, '.pdf'),
          bytes: new Uint8Array(await readFile(pdf)),
          ...(doc.record ? { record: doc.record } : {}),
        })
      }
    } finally {
      await rm(work, { recursive: true, force: true })
    }
  }
  if (output.kind === 'zip') {
    await atomicWriteFile(output.path, Buffer.from(await packageBatch(spec.model, files)))
    return { count: files.length, output: output.path, warnings }
  }
  await mkdir(output.dir, { recursive: true })
  for (const file of files)
    await atomicWriteFile(join(output.dir, file.name), Buffer.from(file.bytes))
  await atomicWriteFile(
    join(output.dir, '_MucLuc.csv'),
    Buffer.from(batchIndexCsv(spec.model, files), 'utf8'),
  )
  return { count: files.length, output: output.dir, warnings }
}

/** Zips existing files with an index of their names (File.Package). */
export async function packageFiles(paths: readonly string[], zipPath: string): Promise<number> {
  const files: { name: string; bytes: Uint8Array }[] = []
  const used = new Set<string>()
  for (const path of paths) {
    let name = path.split(/[\\/]/).pop() ?? 'file'
    for (let n = 2; used.has(name.toLowerCase()); n++)
      name = name.replace(/(\.[^.]*)?$/, ` (${n})$1`)
    used.add(name.toLowerCase())
    files.push({ name, bytes: new Uint8Array(await readFile(path)) })
  }
  const zip = await packageBatch(
    { docId: 'doc_package', schemaVersion: 1, fields: [], collections: [], tables: [], links: [] },
    files,
  )
  await mkdir(dirname(zipPath), { recursive: true })
  await atomicWriteFile(zipPath, Buffer.from(zip))
  return files.length
}

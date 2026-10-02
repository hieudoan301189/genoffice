/**
 * Smart Template filling (P6). A template is a docx whose content controls
 * say what to do:
 *
 * - `dvh:f:<fieldId>`      Smart Field: its text becomes the field value; the
 *                          control and its data binding stay (the model part
 *                          is written with the same values).
 * - `dvh:c:<columnId>`     a column of the current record: inside a repeating
 *                          section the repeated row, elsewhere the batch record.
 * - `dvh:if:<id>`          block shown only when condition `<id>` is true.
 * - `dvh:repeat:<collectionId>[:<conditionId>]`
 *                          block repeated for every record (filtered).
 * - `dvh:repeatrows:<collectionId>[:<conditionId>]`
 *                          a table whose rows holding `dvh:c:` controls repeat
 *                          for every record; rows above and below stay.
 * - `dvh:t:<tableId>`      a DVH.Table, re-rendered from its collection.
 * - `dvh:img:<fieldId>`    a picture whose image becomes the field's image.
 *
 * Section wrappers (`if`, `repeat`, `repeatrows`, `c`) are removed from the
 * output; field, table and image controls stay so the result is still Smart.
 * The output zip is deterministic: the same template and data give the same bytes.
 */
import JSZip from 'jszip'
import {
  fieldText,
  formatCellValue,
  MODEL_NS,
  renderTable,
  setFieldTextInXml,
  type DvhCollection,
  type DvhModel,
  type Scalar,
} from '@genoffice/dvh-model'
import { evaluateExpr, truthy, type ExprScope } from './expr'
import {
  elements,
  repairEmptyCells,
  replaceSpans,
  runsWithText,
  sdtContent,
  sdtTag,
  withSdtContent,
} from './ooxml'
import { applyReleasePolicy, findItemByNamespace, type DvhReleasePolicy } from './release'
import { renderedTableToOoxml } from './table-ooxml'

export interface TemplateRecord {
  readonly collectionId: string
  readonly index: number
  readonly row: readonly Scalar[]
}

export interface TemplateData {
  /** fields, collections, tables and conditions the template reads */
  readonly model: DvhModel
  /** the batch record this document is for (column controls outside sections read it) */
  readonly record?: TemplateRecord
  /** how many documents the batch makes (COUNT in expressions) */
  readonly count?: number
}

export interface TemplateImage {
  readonly bytes: Uint8Array
  /** file extension without the dot: png, jpg, gif… */
  readonly ext: string
}

export interface FillOptions {
  /** number formats of table cells and column values (OS regional locale) */
  readonly locale?: string
  /** loads the picture an image field names (a path or a data URL) */
  readonly images?: (value: Scalar, fieldId: string) => TemplateImage | null
  /** what DVH data the output keeps (default: Smart Data without history) */
  readonly release?: DvhReleasePolicy
}

export interface FillResult {
  readonly bytes: Uint8Array
  readonly warnings: readonly string[]
}

/** a fixed timestamp for every zip entry: same input, same bytes */
const ZIP_DATE = new Date(Date.UTC(2000, 0, 1))

interface RowContext {
  readonly collection: DvhCollection
  readonly row: readonly Scalar[]
  readonly index: number
  readonly count: number
}

class Filler {
  readonly warnings: string[] = []
  readonly imageSwaps: { rId: string; fieldId: string; value: Scalar }[] = []

  constructor(
    readonly data: TemplateData,
    readonly options: FillOptions,
  ) {}

  private get model(): DvhModel {
    return this.data.model
  }

  /** the batch record as a row context (column controls outside sections) */
  recordContext(): RowContext | null {
    const record = this.data.record
    if (!record) return null
    const collection = this.model.collections.find((c) => c.id === record.collectionId)
    return collection
      ? { collection, row: record.row, index: record.index, count: this.data.count ?? 1 }
      : null
  }

  scope(row: RowContext | null): ExprScope {
    const model = this.model
    const columnOf = (key: string) => {
      if (!row) return undefined
      const i = row.collection.columns.findIndex((c) => c.key === key || c.id === key)
      return i >= 0 ? (row.row[i] ?? null) : undefined
    }
    return {
      lookup: (name) => {
        if (name.toUpperCase() === 'INDEX') return row ? row.index + 1 : 1
        if (name.toUpperCase() === 'COUNT') return row ? row.count : (this.data.count ?? 1)
        if (name.startsWith('row.')) return columnOf(name.slice(4))
        const field = model.fields.find((f) => f.name === name || f.id === name)
        if (field) return field.value
        return columnOf(name)
      },
      column: (title) => {
        if (!row) return undefined
        const i = row.collection.columns.findIndex((c) => c.title === title)
        return i >= 0 ? (row.row[i] ?? null) : undefined
      },
    }
  }

  private condition(id: string, scope: ExprScope): boolean {
    const condition = this.model.conditions?.find((c) => c.id === id)
    if (!condition) {
      this.warnings.push(`unknown condition ${id}: its content is kept`)
      return true
    }
    try {
      return truthy(evaluateExpr(condition.expr, scope))
    } catch (error) {
      this.warnings.push(`condition ${condition.name ?? id}: ${(error as Error).message}`)
      return false
    }
  }

  private records(tag: string): RowContext[] {
    const [, , collectionId, conditionId] = tag.split(':')
    const collection = this.model.collections.find((c) => c.id === collectionId)
    if (!collection) {
      this.warnings.push(`unknown collection ${collectionId}: section left empty`)
      return []
    }
    const all = collection.rows.map((row, index) => ({
      collection,
      row,
      index,
      count: collection.rows.length,
    }))
    if (!conditionId) return all
    return all.filter((ctx) => this.condition(conditionId, this.scope(ctx)))
  }

  private columnText(columnId: string, row: RowContext | null): string {
    if (!row) {
      this.warnings.push(`column ${columnId} outside any record: left empty`)
      return ''
    }
    const i = row.collection.columns.findIndex((c) => c.id === columnId || c.key === columnId)
    if (i < 0) {
      this.warnings.push(`unknown column ${columnId}`)
      return ''
    }
    const column = row.collection.columns[i]!
    const value = row.row[i] ?? null
    return typeof value === 'number' && column.numFmt
      ? formatCellValue(value, column.numFmt, this.options.locale)
      : fieldText(value)
  }

  /** Processes one piece of WordprocessingML against a row context. */
  process(xml: string, row: RowContext | null): string {
    const scope = this.scope(row)
    const spans: { start: number; end: number; with: string }[] = []
    for (const sdt of elements(xml, 'w:sdt')) {
      const tag = sdtTag(sdt.xml)
      const content = sdtContent(sdt.xml)
      let replacement: string
      if (tag.startsWith('dvh:repeat:')) {
        replacement = this.records(tag)
          .map((ctx) => this.process(content, ctx))
          .join('')
      } else if (tag.startsWith('dvh:repeatrows:')) {
        replacement = this.repeatRows(content, tag, row)
      } else if (tag.startsWith('dvh:if:')) {
        replacement = this.condition(tag.slice('dvh:if:'.length), scope)
          ? this.process(content, row)
          : ''
      } else if (tag.startsWith('dvh:c:')) {
        replacement = runsWithText(content, this.columnText(tag.slice('dvh:c:'.length), row))
      } else if (tag.startsWith('dvh:f:')) {
        const field = this.model.fields.find((f) => f.id === tag.slice('dvh:f:'.length))
        if (!field) this.warnings.push(`unknown field ${tag.slice('dvh:f:'.length)}`)
        replacement = field
          ? withSdtContent(sdt.xml, runsWithText(content, fieldText(field.value)))
          : sdt.xml
      } else if (tag.startsWith('dvh:t:')) {
        replacement = this.table(sdt.xml, tag.slice('dvh:t:'.length))
      } else if (tag.startsWith('dvh:img:')) {
        replacement = this.image(sdt.xml, tag.slice('dvh:img:'.length))
      } else {
        // a control the engine does not own: its content may still hold DVH controls
        replacement = withSdtContent(sdt.xml, this.process(content, row))
      }
      spans.push({ start: sdt.start, end: sdt.end, with: replacement })
    }
    return replaceSpans(xml, spans)
  }

  private repeatRows(content: string, tag: string, row: RowContext | null): string {
    const table = elements(content, 'w:tbl')[0]
    if (!table) {
      this.warnings.push(`${tag} holds no table: section left as is`)
      return this.process(content, row)
    }
    const rows = elements(table.xml, 'w:tr')
    const template = rows.filter((r) => r.xml.includes('w:val="dvh:c:'))
    if (template.length === 0) return this.process(content, row)
    const first = rows.indexOf(template[0]!)
    const last = rows.indexOf(template.at(-1)!)
    const records = this.records(tag)
    const repeated = records
      .map((ctx) =>
        rows
          .slice(first, last + 1)
          .map((r) => this.process(r.xml, ctx))
          .join(''),
      )
      .join('')
    const before = rows
      .slice(0, first)
      .map((r) => this.process(r.xml, row))
      .join('')
    const after = rows
      .slice(last + 1)
      .map((r) => this.process(r.xml, row))
      .join('')
    const head = table.xml.slice(0, rows[0]!.start)
    const tail = table.xml.slice(rows.at(-1)!.end)
    const newTable = head + before + repeated + after + tail
    return (
      this.process(content.slice(0, table.start), row) +
      newTable +
      this.process(content.slice(table.end), row)
    )
  }

  private table(sdtXml: string, tableId: string): string {
    const table = this.model.tables.find((t) => t.id === tableId)
    if (!table) {
      this.warnings.push(`unknown table ${tableId}: kept as in the template`)
      return sdtXml
    }
    const rendered = renderTable(table, this.model.collections)
    return withSdtContent(sdtXml, renderedTableToOoxml(rendered, table, this.options.locale))
  }

  private image(sdtXml: string, fieldId: string): string {
    const field = this.model.fields.find((f) => f.id === fieldId)
    const rId = /r:embed="([^"]+)"/.exec(sdtXml)?.[1]
    if (!field || !rId) {
      this.warnings.push(`image field ${fieldId}: no field or no picture in the control`)
      return sdtXml
    }
    if (field.value !== null && field.value !== '') {
      this.imageSwaps.push({ rId, fieldId, value: field.value })
    }
    return sdtXml
  }
}

const PART_PATTERN = /^word\/(document|header\d*|footer\d*)\.xml$/

/** Fills a template for one document. */
export async function fillTemplate(
  templateBytes: Uint8Array,
  data: TemplateData,
  options: FillOptions = {},
): Promise<FillResult> {
  const zip = await JSZip.loadAsync(templateBytes)
  const filler = new Filler(data, options)
  const record = filler.recordContext()
  for (const path of Object.keys(zip.files).sort()) {
    if (!PART_PATTERN.test(path)) continue
    const xml = await zip.file(path)!.async('string')
    filler.imageSwaps.length = 0
    const next = repairEmptyCells(filler.process(xml, record))
    zip.file(path, next)
    if (filler.imageSwaps.length > 0) await swapImages(zip, path, filler, options)
  }
  // the model part carries the same values as the controls (Word re-reads it through the bindings)
  const modelPath = await findItemByNamespace(zip, MODEL_NS)
  if (modelPath) {
    let modelXml = await zip.file(modelPath)!.async('string')
    for (const field of data.model.fields) {
      modelXml = setFieldTextInXml(modelXml, field.id, fieldText(field.value)) ?? modelXml
    }
    zip.file(modelPath, modelXml)
  }
  const released = await applyReleasePolicy(
    await deterministicZip(zip),
    options.release ?? { kind: 'none' },
  )
  return {
    bytes: await deterministicZip(await JSZip.loadAsync(released)),
    warnings: filler.warnings,
  }
}

async function swapImages(
  zip: JSZip,
  partPath: string,
  filler: Filler,
  options: FillOptions,
): Promise<void> {
  const relsPath = partPath.replace(/^word\/(.+)$/, 'word/_rels/$1.rels')
  let rels = (await zip.file(relsPath)?.async('string')) ?? ''
  let types = (await zip.file('[Content_Types].xml')?.async('string')) ?? ''
  for (const swap of filler.imageSwaps) {
    const image = options.images?.(swap.value, swap.fieldId) ?? null
    if (!image) {
      filler.warnings.push(
        `image field ${swap.fieldId}: "${fieldText(swap.value)}" could not be loaded`,
      )
      continue
    }
    const ext = image.ext.toLowerCase().replace(/^\./, '')
    const target = `media/dvh_${swap.fieldId}.${ext}`
    zip.file(`word/${target}`, image.bytes)
    rels = rels.replace(
      new RegExp(`(<Relationship\\b[^>]*\\bId="${swap.rId}"[^>]*\\bTarget=")[^"]*(")`),
      `$1${target}$2`,
    )
    if (!new RegExp(`Extension="${ext}"`, 'i').test(types)) {
      const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`
      types = types.replace(
        /(<Types[^>]*>)/,
        `$1<Default Extension="${ext}" ContentType="${mime}"/>`,
      )
    }
  }
  if (rels) zip.file(relsPath, rels)
  if (types) zip.file('[Content_Types].xml', types)
}

/** Re-zips with fixed timestamps and a stable entry order, so equal content means equal bytes. */
export async function deterministicZip(zip: JSZip): Promise<Uint8Array> {
  const out = new JSZip()
  const names = Object.keys(zip.files).filter((n) => !zip.files[n]!.dir)
  // [Content_Types].xml first, as Office writes it; the rest in a stable order
  names.sort((a, b) =>
    a === '[Content_Types].xml' ? -1 : b === '[Content_Types].xml' ? 1 : a < b ? -1 : a > b ? 1 : 0,
  )
  for (const name of names) {
    out.file(name, await zip.file(name)!.async('uint8array'), {
      date: ZIP_DATE,
      createFolders: false,
    })
  }
  return out.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

/**
 * Batch Generator (P6): data → template → filter → preview → documents →
 * package. One document per record of a collection (or one document when no
 * collection is chosen); the record is what column controls outside sections
 * and `row.*` names read. Same template and data, same bytes.
 */
import JSZip from 'jszip'
import { fieldText, type DvhModel, type Scalar } from '@genoffice/dvh-model'
import { evaluateExpr, renderTemplateText, truthy } from './expr'
import { deterministicZip, fillTemplate, type FillOptions, type TemplateRecord } from './fill'

export interface BatchSpec {
  readonly template: Uint8Array
  readonly model: DvhModel
  /** one document per record of this collection; absent: one document */
  readonly collectionId?: string
  /** keeps the records it is true for (expression, see expr.ts) */
  readonly filter?: string
  /** file name rule with `{expression}` holes; `.docx` is appended */
  readonly nameRule: string
  readonly options?: FillOptions
}

export interface BatchPlanItem {
  readonly index: number
  readonly name: string
  readonly record?: TemplateRecord
}

export interface BatchPlan {
  readonly items: readonly BatchPlanItem[]
  /** names produced more than once (a suffix keeps them apart) */
  readonly duplicates: readonly string[]
  readonly warnings: readonly string[]
}

/** characters Windows refuses in file names, and trailing dots/spaces */
export function safeFileName(name: string): string {
  const cleaned = name
    // eslint-disable-next-line no-control-regex -- control characters are not allowed in file names
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/, '')
  const base = cleaned.slice(0, 150) || 'document'
  return /^(con|prn|aux|nul|com\d|lpt\d)$/i.test(base) ? `${base}_` : base
}

/** Records, names and warnings — what the user checks before generating. */
export function planBatch(spec: BatchSpec): BatchPlan {
  const warnings: string[] = []
  const scopeFor = (record: TemplateRecord | undefined, count: number) => {
    const collection = record
      ? spec.model.collections.find((c) => c.id === record.collectionId)
      : undefined
    const columnOf = (key: string): Scalar | undefined => {
      if (!record || !collection) return undefined
      const i = collection.columns.findIndex((c) => c.key === key || c.id === key)
      return i >= 0 ? (record.row[i] ?? null) : undefined
    }
    return {
      lookup: (name: string) => {
        if (name.toUpperCase() === 'INDEX') return record ? record.index + 1 : 1
        if (name.toUpperCase() === 'COUNT') return count
        if (name.startsWith('row.')) return columnOf(name.slice(4))
        const field = spec.model.fields.find((f) => f.name === name || f.id === name)
        return field ? field.value : columnOf(name)
      },
      column: (title: string) => {
        if (!record || !collection) return undefined
        const i = collection.columns.findIndex((c) => c.title === title)
        return i >= 0 ? (record.row[i] ?? null) : undefined
      },
    }
  }

  let records: (TemplateRecord | undefined)[] = [undefined]
  if (spec.collectionId) {
    const collection = spec.model.collections.find((c) => c.id === spec.collectionId)
    if (!collection) throw new Error(`unknown collection ${spec.collectionId}`)
    records = collection.rows.map((row, index) => ({ collectionId: collection.id, index, row }))
  }
  if (spec.filter) {
    const total = records.length
    records = records.filter((record) => {
      try {
        return truthy(evaluateExpr(spec.filter!, scopeFor(record, total)))
      } catch (error) {
        warnings.push(`record ${(record?.index ?? 0) + 1}: ${(error as Error).message}`)
        return false
      }
    })
  }
  const seen = new Map<string, number>()
  const duplicates = new Set<string>()
  const items = records.map((record, i) => {
    let name: string
    try {
      name = safeFileName(renderTemplateText(spec.nameRule, scopeFor(record, records.length)))
    } catch (error) {
      warnings.push(`record ${i + 1}: ${(error as Error).message}`)
      name = `document-${i + 1}`
    }
    const key = name.toLowerCase()
    const n = (seen.get(key) ?? 0) + 1
    seen.set(key, n)
    if (n > 1) {
      duplicates.add(name)
      name = `${name} (${n})`
    }
    return { index: i, name: `${name}.docx`, ...(record ? { record } : {}) }
  })
  return { items, duplicates: [...duplicates], warnings }
}

export interface GeneratedDocument {
  readonly name: string
  readonly bytes: Uint8Array
  readonly record?: TemplateRecord
  readonly warnings: readonly string[]
}

/** Generates every planned document (sequentially: memory stays flat on large batches). */
export async function generateBatch(
  spec: BatchSpec,
  plan: BatchPlan = planBatch(spec),
  onProgress?: (done: number, total: number) => void,
): Promise<GeneratedDocument[]> {
  const out: GeneratedDocument[] = []
  for (const item of plan.items) {
    const result = await fillTemplate(
      spec.template,
      {
        model: spec.model,
        ...(item.record ? { record: item.record } : {}),
        count: plan.items.length,
      },
      spec.options,
    )
    out.push({
      name: item.name,
      bytes: result.bytes,
      ...(item.record ? { record: item.record } : {}),
      warnings: result.warnings,
    })
    onProgress?.(out.length, plan.items.length)
  }
  return out
}

const csvCell = (text: string) => (/[",\n;]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text)

/**
 * The index of a package (mục lục): CSV with a BOM so Excel reads Vietnamese
 * text right; one line per document with its record's columns.
 */
export function batchIndexCsv(
  model: DvhModel,
  documents: readonly { name: string; record?: TemplateRecord }[],
): string {
  const collection = documents[0]?.record
    ? model.collections.find((c) => c.id === documents[0]!.record!.collectionId)
    : undefined
  const header = ['STT', 'File', ...(collection?.columns.map((c) => c.title) ?? [])]
  const lines = documents.map((doc, i) => [
    String(i + 1),
    doc.name,
    ...(doc.record ? doc.record.row.map((v) => fieldText(v ?? null)) : []),
  ])
  return '﻿' + [header, ...lines].map((line) => line.map(csvCell).join(',')).join('\r\n') + '\r\n'
}

/** A zip of the documents plus `_MucLuc.csv`, deterministic like the documents themselves. */
export async function packageBatch(
  model: DvhModel,
  documents: readonly { name: string; bytes: Uint8Array; record?: TemplateRecord }[],
  indexName = '_MucLuc.csv',
): Promise<Uint8Array> {
  const zip = new JSZip()
  for (const doc of documents) zip.file(doc.name, doc.bytes)
  zip.file(indexName, batchIndexCsv(model, documents))
  return deterministicZip(zip)
}

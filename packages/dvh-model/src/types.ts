/**
 * The DVH Smart Data model (ADR D1, D10). The core knows fields, collections
 * and links with user-defined schemas; business vocabularies (QLCL, BOQ...)
 * are schema packs on top, never part of the core.
 */

import { z } from 'zod'

/** `image`: the value is a path or data URL of a picture (Smart Template image fields, P6) */
export const fieldTypeSchema = z.enum(['text', 'number', 'date', 'boolean', 'enum', 'image'])
export type FieldType = z.infer<typeof fieldTypeSchema>

export const scalarSchema = z.union([z.string(), z.number(), z.boolean(), z.null()])
export type Scalar = z.infer<typeof scalarSchema>

export const dvhFieldSchema = z
  .object({
    id: z.string().min(1),
    /** alias shown to people (`Project.Name`); renaming never breaks links */
    name: z.string().min(1),
    type: fieldTypeSchema,
    value: scalarSchema,
    enumValues: z.array(z.string()).optional(),
    access: z.enum(['read', 'readwrite']),
    /**
     * bumped by every DVH write of the value (P3 write-back): a linked document
     * compares it with the revision it last synced to detect conflicts.
     * Writers that know nothing of DVH (Excel) leave it alone, so the value is
     * compared as well.
     */
    rev: z.number().int().nonnegative().optional(),
  })
  .strict()
export type DvhField = z.infer<typeof dvhFieldSchema>

const hexColorSchema = z.string().regex(/^#[0-9A-Fa-f]{6}$/)
export const cellAlignSchema = z.enum(['left', 'center', 'right'])

/**
 * The cell formatting a DVH table carries between Sheets and Docs: the part
 * both can render. Colors are document data (`#RRGGBB`), never theme tokens.
 */
export const dvhCellStyleSchema = z
  .object({
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    color: hexColorSchema.optional(),
    fill: hexColorSchema.optional(),
    align: cellAlignSchema.optional(),
    /** points */
    fontSize: z.number().positive().optional(),
    border: z
      .object({ color: hexColorSchema, width: z.enum(['thin', 'medium']) })
      .strict()
      .optional(),
  })
  .strict()
export type DvhCellStyle = z.infer<typeof dvhCellStyleSchema>

export const dvhColumnSchema = z
  .object({
    id: z.string().min(1),
    key: z.string().min(1),
    title: z.string(),
    type: fieldTypeSchema,
    numFmt: z.string().optional(),
    /** source formatting of the body cells and of the title cell (Keep Source Style) */
    style: dvhCellStyleSchema.optional(),
    headerStyle: dvhCellStyleSchema.optional(),
    /** source column width in pixels */
    width: z.number().positive().optional(),
  })
  .strict()
export type DvhColumn = z.infer<typeof dvhColumnSchema>

export const dvhCollectionSchema = z
  .object({
    id: z.string().min(1),
    name: z.string().min(1),
    columns: z.array(dvhColumnSchema),
    rows: z.array(z.array(scalarSchema)),
  })
  .strict()
export type DvhCollection = z.infer<typeof dvhCollectionSchema>

export const tableTotalSchema = z.enum(['none', 'sum', 'count', 'avg', 'min', 'max'])
export type TableTotal = z.infer<typeof tableTotalSchema>

export const dvhTableColumnSchema = z
  .object({
    id: z.string().min(1),
    /** the collection column shown here (by id); absent for embedded data */
    columnId: z.string().min(1).optional(),
    title: z.string(),
    numFmt: z.string().optional(),
    align: cellAlignSchema.optional(),
    /** relative width weight */
    width: z.number().positive().optional(),
    total: tableTotalSchema.optional(),
  })
  .strict()
export type DvhTableColumn = z.infer<typeof dvhTableColumnSchema>

/**
 * A DVH.Table (P2): which data, which columns, and how it looks wherever it is
 * rendered. `source` style keeps the collection's own formatting, `destination`
 * applies the table's style (or the defaults) to every cell.
 */
export const dvhTableSchema = z
  .object({
    id: z.string().min(1),
    name: z.string().min(1),
    source: z.union([
      z.object({ collectionId: z.string().min(1) }).strict(),
      z.object({ rows: z.array(z.array(scalarSchema)) }).strict(),
    ]),
    columns: z.array(dvhTableColumnSchema).min(1),
    /**
     * the table shows every column of its collection, in the collection's order
     * and with its current titles (columns added, removed or renamed at the
     * source follow on refresh); `columns` then only keeps per-column settings
     */
    autoColumns: z.boolean().optional(),
    /** header rows above the column titles; each cell spans consecutive columns */
    headerGroups: z
      .array(z.array(z.object({ title: z.string(), span: z.number().int().min(1) }).strict()))
      .optional(),
    totalRow: z.object({ label: z.string() }).strict().optional(),
    style: z
      .object({
        mode: z.enum(['source', 'destination']),
        header: dvhCellStyleSchema.optional(),
        body: dvhCellStyleSchema.optional(),
        total: dvhCellStyleSchema.optional(),
        /** fill of every second body row */
        bandFill: hexColorSchema.optional(),
      })
      .strict(),
    layout: z
      .object({
        repeatHeader: z.boolean(),
        keepRowsTogether: z.boolean(),
        widths: z.enum(['auto', 'page']),
      })
      .strict(),
    /**
     * What this document last rendered (hash of the cells as its renderer reads
     * them back): a refresh that finds other content asks before overwriting.
     */
    lastRender: z.object({ hash: z.string(), at: z.string() }).strict().optional(),
  })
  .strict()
export type DvhTable = z.infer<typeof dvhTableSchema>

/** A field as both sides last agreed on it: the source revision and the text. */
export const fieldBaseSchema = z
  .object({ rev: z.number().int().nonnegative(), text: z.string() })
  .strict()
export type FieldBase = z.infer<typeof fieldBaseSchema>

export const dvhLinkSchema = z
  .object({
    id: z.string().min(1),
    /** where the values come from: another DVH document and an object in it */
    source: z
      .object({
        docId: z.string().min(1),
        relPath: z.string().optional(),
        /** last known absolute path, the fallback when relPath no longer resolves */
        path: z.string().optional(),
        objectId: z.string().min(1),
      })
      .strict(),
    /** objects in this document fed by the source */
    targets: z.array(z.string().min(1)),
    update: z.enum(['manual', 'onOpen', 'auto']),
    lastSync: z
      .object({
        revision: z.number().int(),
        hash: z.string(),
        at: z.string(),
        /**
         * per field: the source revision and text this document last agreed on
         * (the common base of a three-way compare); absent in links made before P3
         */
        fields: z.record(z.string(), fieldBaseSchema).optional(),
      })
      .strict()
      .optional(),
  })
  .strict()
export type DvhLink = z.infer<typeof dvhLinkSchema>

/**
 * A Smart Template condition (P6): the block control tagged `dvh:if:<id>`
 * shows its content only when `expr` is true; a repeating section tagged
 * `dvh:repeat:<collectionId>:<id>` keeps only the records it is true for.
 */
export const dvhConditionSchema = z
  .object({
    id: z.string().min(1),
    expr: z.string().min(1).max(2000),
    /** what the author called it */
    name: z.string().optional(),
  })
  .strict()
export type DvhCondition = z.infer<typeof dvhConditionSchema>

export const dvhModelSchema = z
  .object({
    docId: z.string().min(1),
    schemaVersion: z.literal(1),
    fields: z.array(dvhFieldSchema),
    collections: z.array(dvhCollectionSchema),
    /** P2; absent in P1 files */
    tables: z.array(dvhTableSchema).default([]),
    links: z.array(dvhLinkSchema),
    /** P6; absent in documents that are not Smart Templates (keeps their hashes) */
    conditions: z.array(dvhConditionSchema).optional(),
  })
  .strict()
export type DvhModel = z.infer<typeof dvhModelSchema>

export function emptyModel(docId: string): DvhModel {
  return { docId, schemaVersion: 1, fields: [], collections: [], tables: [], links: [] }
}

/** A field value as text (what a bound content control or cell shows). */
export function fieldText(value: Scalar): string {
  if (value === null) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return String(value)
}

/** Reads a field's text back as its declared type. */
export function fieldValueFromText(type: FieldType, text: string): Scalar {
  if (text === '') return null
  if (type === 'number') {
    const n = Number(text.replace(/\s/g, ''))
    return Number.isFinite(n) ? n : text
  }
  if (type === 'boolean') {
    const t = text.trim().toUpperCase()
    if (t === 'TRUE') return true
    if (t === 'FALSE') return false
  }
  return text
}

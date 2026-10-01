/**
 * The DVH Smart Data model (ADR D1, D10). The core knows fields, collections
 * and links with user-defined schemas; business vocabularies (QLCL, BOQ...)
 * are schema packs on top, never part of the core.
 */

import { z } from 'zod'

export const fieldTypeSchema = z.enum(['text', 'number', 'date', 'boolean', 'enum'])
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
  })
  .strict()
export type DvhField = z.infer<typeof dvhFieldSchema>

export const dvhColumnSchema = z
  .object({
    id: z.string().min(1),
    key: z.string().min(1),
    title: z.string(),
    type: fieldTypeSchema,
    numFmt: z.string().optional(),
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

export const dvhLinkSchema = z
  .object({
    id: z.string().min(1),
    /** where the values come from: another DVH document and an object in it */
    source: z
      .object({
        docId: z.string().min(1),
        relPath: z.string().optional(),
        objectId: z.string().min(1),
      })
      .strict(),
    /** objects in this document fed by the source */
    targets: z.array(z.string().min(1)),
    update: z.enum(['manual', 'onOpen', 'auto']),
    lastSync: z
      .object({ revision: z.number().int(), hash: z.string(), at: z.string() })
      .strict()
      .optional(),
  })
  .strict()
export type DvhLink = z.infer<typeof dvhLinkSchema>

export const dvhModelSchema = z
  .object({
    docId: z.string().min(1),
    schemaVersion: z.literal(1),
    fields: z.array(dvhFieldSchema),
    collections: z.array(dvhCollectionSchema),
    links: z.array(dvhLinkSchema),
  })
  .strict()
export type DvhModel = z.infer<typeof dvhModelSchema>

export function emptyModel(docId: string): DvhModel {
  return { docId, schemaVersion: 1, fields: [], collections: [], links: [] }
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

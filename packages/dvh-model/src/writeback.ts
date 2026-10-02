/**
 * Field write-back (P3, Read/Write links): a document edits a linked field and
 * writes it back to the source workbook. Only single fields travel this way;
 * tables stay one-way (plan decision H).
 *
 * Conflicts are found per field with a three-way compare: the base is what
 * both sides last agreed on (source revision + text, kept in the link), the
 * local side is the document's text, the remote side is the source as it is
 * now. The writer re-checks the base right before it writes, so a source
 * edited in between is never overwritten silently.
 */

import { z } from 'zod'

import { fieldBaseSchema, fieldText, scalarSchema, type DvhField, type FieldBase } from './types'

/** document renderer → main (invoke) */
export const DVH_WRITE_FIELDS_CHANNEL = 'docs:dvh-write-fields'
/** main → the Sheets tab holding the source workbook */
export const DVH_SHEETS_SET_FIELDS_CHANNEL = 'sheets:dvh-set-fields'
/** Sheets tab → main: the answer to one set-fields request */
export const DVH_SHEETS_SET_FIELDS_RESULT_CHANNEL = 'dvh:set-fields-result'

export const fieldWriteSchema = z
  .object({
    fieldId: z.string().min(1).max(128),
    value: scalarSchema,
    /** the base the document edited from; null writes without a check */
    expected: fieldBaseSchema.nullable(),
    /** keep this side: write even when the source moved on */
    force: z.boolean(),
  })
  .strict()
export type FieldWrite = z.infer<typeof fieldWriteSchema>

export const fieldWriteStatusSchema = z.enum([
  /** the source now holds the value */
  'written',
  /** the source changed since the base: nothing written */
  'conflict',
  /** the source has no such field (or its cell is gone) */
  'missing',
  /** the bound cell holds a formula; its value is computed, not written */
  'formula',
])
export type FieldWriteStatus = z.infer<typeof fieldWriteStatusSchema>

export const fieldWriteResultSchema = z
  .object({
    fieldId: z.string().min(1),
    status: fieldWriteStatusSchema,
    /** the source field after the attempt (the new base when written) */
    current: fieldBaseSchema.optional(),
  })
  .strict()
export type FieldWriteResult = z.infer<typeof fieldWriteResultSchema>

export const writeFieldsRequestSchema = z
  .object({
    /** the source workbook's DVH docId */
    docId: z.string().min(1),
    /** candidate paths of the workbook on disk, in order */
    paths: z.array(z.string().min(1).max(4096)).max(16),
    writes: z.array(fieldWriteSchema).min(1).max(256),
    /** the writing document, for the change set's actor */
    origin: z.object({ docId: z.string().min(1), name: z.string().max(512).optional() }).strict(),
  })
  .strict()
export type WriteFieldsRequest = z.infer<typeof writeFieldsRequestSchema>

export const writeFieldsResultSchema = z
  .object({
    /** live: the workbook is open in Sheets; file: the xlsx on disk was patched */
    via: z.enum(['live', 'file', 'none']),
    path: z.string().nullable(),
    results: z.array(fieldWriteResultSchema),
    error: z.string().optional(),
  })
  .strict()
export type WriteFieldsResult = z.infer<typeof writeFieldsResultSchema>

export const sheetsSetFieldsRequestSchema = z
  .object({
    requestId: z.string().min(1).max(64),
    docId: z.string().min(1),
    writes: z.array(fieldWriteSchema).min(1).max(256),
    origin: writeFieldsRequestSchema.shape.origin,
  })
  .strict()
export type SheetsSetFieldsRequest = z.infer<typeof sheetsSetFieldsRequestSchema>

export const sheetsSetFieldsReplySchema = z
  .object({
    requestId: z.string().min(1).max(64),
    /** false: this tab does not hold that workbook */
    handled: z.boolean(),
    path: z.string().nullable(),
    results: z.array(fieldWriteResultSchema),
  })
  .strict()
export type SheetsSetFieldsReply = z.infer<typeof sheetsSetFieldsReplySchema>

/** A field as a sync base: its revision (0 when it never had one) and its text. */
export function fieldBaseOf(field: Pick<DvhField, 'rev' | 'value'>): FieldBase {
  return { rev: field.rev ?? 0, text: fieldText(field.value) }
}

/** True when the source moved away from the base (another revision or other text). */
export function sourceMoved(base: FieldBase | undefined, current: FieldBase): boolean {
  return base !== undefined && (base.rev !== current.rev || base.text !== current.text)
}

/**
 * What a writer does with one write against the field as it stands: write it,
 * or report a conflict. A write equal to the source is still `written` (the
 * two sides agree, so the base moves on).
 */
export function checkFieldWrite(write: FieldWrite, current: FieldBase): 'write' | 'conflict' {
  if (write.force || write.expected === null) return 'write'
  if (current.text === fieldText(write.value)) return 'write'
  return sourceMoved(write.expected, current) ? 'conflict' : 'write'
}

export type ThreeWay = 'same' | 'local' | 'source' | 'conflict'

/**
 * Three-way state of one field: `local` = only the document changed (write it
 * back), `source` = only the source changed (pull it), `conflict` = both, to
 * different texts. Without a base the source wins, as links made before P3 did.
 */
export function threeWay(base: FieldBase | undefined, local: string, source: FieldBase): ThreeWay {
  if (local === source.text) return 'same'
  if (!base) return 'source'
  const localChanged = local !== base.text
  const sourceChanged = sourceMoved(base, source)
  if (localChanged && sourceChanged) return 'conflict'
  if (localChanged) return 'local'
  return 'source'
}

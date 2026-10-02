/**
 * Schema packs (P10, ADR D10): the core knows no business vocabulary. A pack
 * installed into a project defines object types (their fields and the
 * relations between them), validation rules written as safe expressions (P6)
 * and the templates that come with it. Construction/QLCL is one such pack,
 * shipped as data, never as code.
 */

import { z } from 'zod'
import { fieldTypeSchema } from '@genoffice/dvh-model'
import { exprError } from '@genoffice/dvh-template'

export const KEY = /^[A-Za-z_][A-Za-z0-9_]*$/
const key = z.string().regex(KEY, 'a key is letters, digits and _')
const expr = z
  .string()
  .min(1)
  .max(2000)
  .refine((source) => exprError(source) === null, {
    error: (issue) => `bad expression: ${exprError(String(issue.input))}`,
  })

export const packFieldSchema = z
  .object({
    key,
    label: z.string().optional(),
    /** `ref`: the id of another object, of type `ref` */
    type: z.union([fieldTypeSchema, z.literal('ref')]),
    ref: z.string().optional(),
    required: z.boolean().optional(),
    enumValues: z.array(z.string()).optional(),
  })
  .strict()
  .refine((f) => (f.type === 'ref') === (f.ref !== undefined), {
    error: 'a ref field names its target type (and only a ref field does)',
  })
export type PackField = z.infer<typeof packFieldSchema>

export const packTypeSchema = z
  .object({
    /** the type name, also the prefix of its field names (`Project.Name`) */
    name: key,
    label: z.string().optional(),
    /** one object per project (shown as fields), not a list (shown as a collection) */
    single: z.boolean().optional(),
    fields: z.array(packFieldSchema).min(1),
  })
  .strict()
export type PackType = z.infer<typeof packTypeSchema>

export const packRuleSchema = z
  .object({
    id: key,
    /** the type the rule checks, per object */
    type: key,
    /** only objects for which this holds are checked */
    when: expr.optional(),
    /** must be true */
    expr,
    message: z.string().min(1),
    severity: z.enum(['error', 'warning']),
  })
  .strict()
export type PackRule = z.infer<typeof packRuleSchema>

export const schemaPackSchema = z
  .object({
    format: z.literal('dvh-schema-pack'),
    version: z.literal(1),
    id: z.string().regex(/^[a-z0-9][a-z0-9-]*$/, 'a pack id is lowercase letters, digits and -'),
    name: z.string().min(1),
    packVersion: z.string().min(1),
    description: z.string().optional(),
    types: z.array(packTypeSchema).min(1),
    rules: z.array(packRuleSchema).default([]),
    /** templates that come with the pack (files next to the pack) */
    templates: z
      .array(
        z
          .object({ name: z.string(), file: z.string(), description: z.string().optional() })
          .strict(),
      )
      .default([]),
  })
  .strict()
  .superRefine((pack, ctx) => {
    const names = new Set<string>()
    for (const type of pack.types) {
      if (names.has(type.name)) ctx.addIssue({ code: 'custom', message: `type ${type.name} twice` })
      names.add(type.name)
      const keys = new Set<string>()
      for (const field of type.fields) {
        if (field.key === 'id')
          ctx.addIssue({
            code: 'custom',
            message: `${type.name}: "id" is reserved for the object id`,
          })
        if (keys.has(field.key))
          ctx.addIssue({ code: 'custom', message: `${type.name}.${field.key} twice` })
        keys.add(field.key)
      }
    }
    for (const type of pack.types) {
      for (const field of type.fields) {
        if (field.ref && !names.has(field.ref))
          ctx.addIssue({
            code: 'custom',
            message: `${type.name}.${field.key} refers to unknown type ${field.ref}`,
          })
      }
    }
    for (const rule of pack.rules) {
      if (!names.has(rule.type))
        ctx.addIssue({
          code: 'custom',
          message: `rule ${rule.id} checks unknown type ${rule.type}`,
        })
    }
  })
export type SchemaPack = z.infer<typeof schemaPackSchema>

/** Parses a pack file; one readable error for the first problem. */
export function parseSchemaPack(value: unknown): SchemaPack {
  const parsed = schemaPackSchema.safeParse(value)
  if (!parsed.success) {
    const issue = parsed.error.issues[0]!
    throw new Error(`invalid schema pack at ${issue.path.join('.') || '(root)'}: ${issue.message}`)
  }
  return parsed.data
}

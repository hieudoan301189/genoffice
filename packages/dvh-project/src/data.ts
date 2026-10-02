/**
 * The project data store (P10): the objects of one project, typed by the
 * installed schema packs. Every value carries a revision (as Smart Data
 * fields do since P3), so documents linked to project objects and exchange
 * partners can detect conflicts; every change is a change set (P5).
 *
 * Pure data and functions: the project-store package keeps the JSON on disk.
 */

import { z } from 'zod'
import {
  changeSetSchema,
  fieldText,
  fieldValueFromText,
  scalarSchema,
  type ChangeSet,
  type Scalar,
} from '@genoffice/dvh-model'
import { evaluateExpr, truthy } from '@genoffice/dvh-template'
import { schemaPackSchema, type PackField, type PackType, type SchemaPack } from './pack'

export const projectObjectSchema = z
  .object({
    id: z.string().regex(/^[A-Za-z0-9_-]{1,64}$/),
    type: z.string().min(1),
    /** bumped by every change of the object */
    rev: z.number().int().nonnegative(),
    values: z.record(z.string(), scalarSchema),
    /** per value: its revision (the base of a write-back or a sync) */
    revs: z.record(z.string(), z.number().int().nonnegative()),
  })
  .strict()
export type ProjectObject = z.infer<typeof projectObjectSchema>

/** A value or an object both sides changed since the last sync: it waits for the user. */
export const syncConflictSchema = z
  .object({
    objectId: z.string().min(1),
    type: z.string().min(1),
    /** the value's key; absent for a deletion conflict */
    key: z.string().optional(),
    kind: z.enum(['value', 'deleted-there', 'deleted-here']),
    base: z.string().nullable(),
    local: z.string().nullable(),
    remote: z.string().nullable(),
    /** the remote object's values (deletion conflicts) */
    remoteValues: z.record(z.string(), z.string()).optional(),
  })
  .strict()
export type SyncConflict = z.infer<typeof syncConflictSchema>

/** What an exchange partner and this project last agreed on: per object, per key, the text. */
export const syncBaseSchema = z
  .object({
    at: z.string(),
    objects: z.record(
      z.string(),
      z.object({ type: z.string(), values: z.record(z.string(), z.string()) }).strict(),
    ),
    /** conflicts of the last sync still waiting for the user (kept across restarts) */
    conflicts: z.array(syncConflictSchema).optional(),
  })
  .strict()
export type SyncBase = z.infer<typeof syncBaseSchema>

export const projectDataSchema = z
  .object({
    format: z.literal('dvh-project-data'),
    version: z.literal(1),
    projectId: z.string().min(1),
    /** the DVH docId documents link to (`proj_<projectId>`) */
    docId: z.string().min(1),
    /** bumped by every change */
    revision: z.number().int().nonnegative(),
    packs: z.array(schemaPackSchema),
    objects: z.array(projectObjectSchema),
    /** the newest change sets (capped) */
    history: z.array(changeSetSchema),
    /** per exchange partner (a QLCL file, by path) */
    sync: z.record(z.string(), syncBaseSchema),
  })
  .strict()
export type ProjectData = z.infer<typeof projectDataSchema>

export const HISTORY_CAP = 500

export function emptyProjectData(projectId: string): ProjectData {
  return {
    format: 'dvh-project-data',
    version: 1,
    projectId,
    docId: `proj_${projectId.replace(/[^A-Za-z0-9_-]/g, '_')}`,
    revision: 0,
    packs: [],
    objects: [],
    history: [],
    sync: {},
  }
}

/** Reads stored project data; a store that is not project data starts empty. */
export function parseProjectData(value: unknown, projectId: string): ProjectData {
  const parsed = projectDataSchema.safeParse(value)
  return parsed.success && parsed.data.projectId === projectId
    ? parsed.data
    : emptyProjectData(projectId)
}

const ALPHABET = 'abcdefghijklmnopqrstuvwxyz234567'
export function newObjectId(): string {
  const bytes = new Uint8Array(16)
  globalThis.crypto.getRandomValues(bytes)
  let body = ''
  for (const b of bytes) body += ALPHABET[b & 31]
  return `o_${body}`
}

export interface ChangeContext {
  readonly source: ChangeSet['source']
  readonly action: string
  readonly at?: string
  readonly txId?: string
  readonly actor?: ChangeSet['actor']
}

let seq = 0
function record(data: ProjectData, ctx: ChangeContext, changes: ChangeSet['changes']): ProjectData {
  if (changes.length === 0) return data
  const at = ctx.at ?? new Date().toISOString()
  const id = `cs_${Date.now().toString(36)}${(seq++).toString(36).padStart(4, '0')}`
  const changeSet: ChangeSet = {
    id,
    txId: ctx.txId ?? id.replace(/^cs_/, 'tx_'),
    docId: data.docId,
    at,
    source: ctx.source,
    ...(ctx.actor ? { actor: ctx.actor } : {}),
    action: ctx.action,
    changes,
  }
  return {
    ...data,
    revision: data.revision + 1,
    history: [...data.history, changeSet].slice(-HISTORY_CAP),
  }
}

// ---------------- types ----------------

/** Every type of the installed packs. */
export function projectTypes(data: ProjectData): PackType[] {
  return data.packs.flatMap((p) => p.types)
}

export function typeOf(data: ProjectData, name: string): PackType | undefined {
  return projectTypes(data).find((t) => t.name === name)
}

/** Installs (or upgrades) a pack; type names must stay unique across packs. */
export function installPack(data: ProjectData, pack: SchemaPack, ctx: ChangeContext): ProjectData {
  const checked = schemaPackSchema.parse(pack)
  const others = data.packs.filter((p) => p.id !== checked.id)
  const taken = new Set(others.flatMap((p) => p.types.map((t) => t.name)))
  const clash = checked.types.find((t) => taken.has(t.name))
  if (clash) throw new Error(`type ${clash.name} already comes from another pack`)
  const before = data.packs.find((p) => p.id === checked.id)?.packVersion ?? null
  return record({ ...data, packs: [...others, checked] }, ctx, [
    { objectId: `pack:${checked.id}`, path: 'packVersion', before, after: checked.packVersion },
  ])
}

/** Removes a pack; its objects stay (shown as untyped until a pack defines them again). */
export function removePack(data: ProjectData, packId: string, ctx: ChangeContext): ProjectData {
  const pack = data.packs.find((p) => p.id === packId)
  if (!pack) return data
  return record({ ...data, packs: data.packs.filter((p) => p.id !== packId) }, ctx, [
    { objectId: `pack:${packId}`, path: 'packVersion', before: pack.packVersion, after: null },
  ])
}

// ---------------- objects ----------------

/** A value as the field's type wants it (text from a cell or a form). */
export function coerceValue(field: Pick<PackField, 'type'>, value: Scalar): Scalar {
  if (value === null || field.type === 'ref') return value
  if (typeof value === 'string')
    return fieldValueFromText(field.type === 'image' ? 'text' : field.type, value)
  return value
}

export function createObject(
  data: ProjectData,
  type: string,
  values: Record<string, Scalar>,
  ctx: ChangeContext,
  id = newObjectId(),
): { data: ProjectData; object: ProjectObject } {
  const def = typeOf(data, type)
  if (!def) throw new Error(`unknown type ${type}`)
  if (data.objects.some((o) => o.id === id)) throw new Error(`object ${id} exists`)
  if (def.single && data.objects.some((o) => o.type === type))
    throw new Error(`${type} is one per project`)
  const clean: Record<string, Scalar> = {}
  for (const [k, v] of Object.entries(values)) {
    const field = def.fields.find((f) => f.key === k)
    if (!field) throw new Error(`${type} has no field ${k}`)
    clean[k] = coerceValue(field, v)
  }
  const object: ProjectObject = {
    id,
    type,
    rev: 1,
    values: clean,
    revs: Object.fromEntries(Object.keys(clean).map((k) => [k, 1])),
  }
  const next = record({ ...data, objects: [...data.objects, object] }, ctx, [
    { objectId: id, path: '', before: null, after: { type, values: clean } },
  ])
  return { data: next, object }
}

/** Sets values of one object; each changed value gets a new revision. */
export function setValues(
  data: ProjectData,
  objectId: string,
  values: Record<string, Scalar>,
  ctx: ChangeContext,
): ProjectData {
  const object = data.objects.find((o) => o.id === objectId)
  if (!object) throw new Error(`no object ${objectId}`)
  const def = typeOf(data, object.type)
  const changes: ChangeSet['changes'] = []
  const nextValues = { ...object.values }
  const nextRevs = { ...object.revs }
  for (const [k, raw] of Object.entries(values)) {
    const field = def?.fields.find((f) => f.key === k)
    if (def && !field) throw new Error(`${object.type} has no field ${k}`)
    const value = field ? coerceValue(field, raw) : raw
    const before = object.values[k] ?? null
    const same =
      before === value ||
      (before !== null && value !== null && fieldText(before) === fieldText(value))
    if (same) continue
    nextValues[k] = value
    nextRevs[k] = (object.revs[k] ?? 0) + 1
    changes.push({ objectId, path: `values.${k}`, before, after: value })
  }
  if (changes.length === 0) return data
  const updated: ProjectObject = {
    ...object,
    rev: object.rev + 1,
    values: nextValues,
    revs: nextRevs,
  }
  return record(
    { ...data, objects: data.objects.map((o) => (o.id === objectId ? updated : o)) },
    ctx,
    changes,
  )
}

export function deleteObject(data: ProjectData, objectId: string, ctx: ChangeContext): ProjectData {
  const object = data.objects.find((o) => o.id === objectId)
  if (!object) return data
  return record({ ...data, objects: data.objects.filter((o) => o.id !== objectId) }, ctx, [
    { objectId, path: '', before: { type: object.type, values: object.values }, after: null },
  ])
}

// ---------------- validation ----------------

export interface ProjectIssue {
  readonly objectId: string
  readonly field?: string
  readonly rule?: string
  readonly severity: 'error' | 'warning'
  readonly message: string
}

/** Checks every object against its type and the packs' rules. */
export function validateProject(data: ProjectData): ProjectIssue[] {
  const issues: ProjectIssue[] = []
  const byId = new Map(data.objects.map((o) => [o.id, o]))
  const types = new Map(projectTypes(data).map((t) => [t.name, t]))
  const singles = new Map<string, number>()
  for (const object of data.objects) {
    const def = types.get(object.type)
    if (!def) {
      issues.push({
        objectId: object.id,
        severity: 'warning',
        message: `no installed pack defines ${object.type}`,
      })
      continue
    }
    if (def.single) singles.set(def.name, (singles.get(def.name) ?? 0) + 1)
    for (const field of def.fields) {
      const value = object.values[field.key] ?? null
      const at = { objectId: object.id, field: field.key }
      if (value === null || value === '') {
        if (field.required)
          issues.push({ ...at, severity: 'error', message: `${def.name}.${field.key} is required` })
        continue
      }
      if (field.type === 'number' && typeof value !== 'number')
        issues.push({
          ...at,
          severity: 'error',
          message: `${def.name}.${field.key} must be a number`,
        })
      if (field.type === 'enum' && field.enumValues && !field.enumValues.includes(String(value)))
        issues.push({
          ...at,
          severity: 'error',
          message: `${def.name}.${field.key} must be one of ${field.enumValues.join(', ')}`,
        })
      if (field.type === 'ref') {
        const target = byId.get(String(value))
        if (!target || target.type !== field.ref)
          issues.push({
            ...at,
            severity: 'error',
            message: `${def.name}.${field.key} refers to a missing ${field.ref}`,
          })
      }
    }
  }
  for (const [name, count] of singles) {
    if (count > 1)
      issues.push({
        objectId: '',
        severity: 'error',
        message: `${name} must be one per project (found ${count})`,
      })
  }
  for (const pack of data.packs) {
    for (const rule of pack.rules) {
      for (const object of data.objects.filter((o) => o.type === rule.type)) {
        const scope = objectScope(data, object, byId, types)
        try {
          if (rule.when && !truthy(evaluateExpr(rule.when, scope))) continue
          if (!truthy(evaluateExpr(rule.expr, scope)))
            issues.push({
              objectId: object.id,
              rule: rule.id,
              severity: rule.severity,
              message: rule.message,
            })
        } catch (error) {
          issues.push({
            objectId: object.id,
            rule: rule.id,
            severity: 'warning',
            message: `rule ${rule.id} could not be checked: ${error instanceof Error ? error.message : String(error)}`,
          })
        }
      }
    }
  }
  return issues
}

/** Names an expression sees for one object: its keys, `id`, and `ref.key` through references. */
function objectScope(
  data: ProjectData,
  object: ProjectObject,
  byId: Map<string, ProjectObject>,
  types: Map<string, PackType>,
) {
  const def = types.get(object.type)
  return {
    lookup: (name: string): Scalar | undefined => {
      const [head, ...rest] = name.split('.')
      if (rest.length === 0) {
        if (head === 'id') return object.id
        if (def?.fields.some((f) => f.key === head)) return object.values[head!] ?? null
        return undefined
      }
      const field = def?.fields.find((f) => f.key === head && f.type === 'ref')
      if (!field || rest.length !== 1) return undefined
      const target = byId.get(String(object.values[head!] ?? ''))
      if (!target) return null
      if (rest[0] === 'id') return target.id
      return types.get(target.type)?.fields.some((f) => f.key === rest[0])
        ? (target.values[rest[0]!] ?? null)
        : undefined
    },
  }
}

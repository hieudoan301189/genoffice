/**
 * Two-way exchange with QLCL-DVH at project level (P10). Both sides speak the
 * P6 exchange model (the QLCL workbook or the dvh-exchange JSON); objects are
 * matched by id (the `ID` column, or the type for one-per-project objects).
 *
 * Every value is compared three ways, as links do since P3: the base is what
 * the two sides agreed on at the last sync with this partner. Only one side
 * changed → that side wins; both changed to different texts → a conflict,
 * written to neither side until the user picks. Deletions are three-way too:
 * an object deleted on one side and unchanged on the other is deleted; deleted
 * on one side and edited on the other is a conflict.
 *
 * Without a base (the first sync) differing values are conflicts unless one
 * side is empty: two independent sources are never merged by guess.
 */

import { fieldText, type DvhModel, type Scalar } from '@genoffice/dvh-model'
import {
  coerceValue,
  createObject,
  deleteObject,
  projectTypes,
  setValues,
  typeOf,
  type ChangeContext,
  type ProjectData,
  type SyncBase,
  type SyncConflict,
} from './data'
import type { PackType } from './pack'
import { projectSourceModel } from './source'

export type { SyncConflict }

export interface SyncResult {
  readonly data: ProjectData
  /** what to write back to the partner (conflicts keep the partner's own values) */
  readonly outgoing: DvhModel
  readonly conflicts: readonly SyncConflict[]
  readonly stats: {
    readonly pulled: number
    readonly pushed: number
    readonly created: number
    readonly deleted: number
  }
}

type Texts = Record<string, string>
interface RemoteObject {
  readonly id: string | null
  readonly type: string
  readonly values: Texts
}

const ID_TITLES = new Set(['id', 'ID', 'Id', 'Mã ID'])

/** The partner's objects, read from its exchange model by type and id. */
export function remoteObjects(data: ProjectData, incoming: DvhModel): RemoteObject[] {
  const out: RemoteObject[] = []
  for (const type of projectTypes(data)) {
    if (type.single) {
      const values: Texts = {}
      for (const field of type.fields) {
        const found = incoming.fields.find((f) => f.name === `${type.name}.${field.key}`)
        if (found) values[field.key] = fieldText(found.value)
      }
      if (Object.keys(values).length === 0) continue
      const local = data.objects.find((o) => o.type === type.name)
      out.push({ id: local?.id ?? null, type: type.name, values })
      continue
    }
    const collection = incoming.collections.find(
      (c) => c.name === type.name || (type.label !== undefined && c.name === type.label),
    )
    if (!collection) continue
    const idIndex = collection.columns.findIndex((c) => c.key === 'id' || ID_TITLES.has(c.title))
    const columnOf = (field: PackType['fields'][number]) =>
      collection.columns.findIndex(
        (c) =>
          c.key === field.key ||
          c.title === field.key ||
          (field.label !== undefined && c.title === field.label),
      )
    const indexes = type.fields.map((f) => [f.key, columnOf(f)] as const)
    for (const row of collection.rows) {
      const id = idIndex >= 0 ? fieldText(row[idIndex] ?? null).trim() : ''
      const values: Texts = {}
      for (const [key, index] of indexes)
        if (index >= 0) values[key] = fieldText(row[index] ?? null)
      if (!id && Object.values(values).every((v) => v === '')) continue
      out.push({ id: id || null, type: type.name, values })
    }
  }
  return out
}

const textsOf = (values: Record<string, Scalar>, keys: readonly string[]): Texts =>
  Object.fromEntries(keys.map((k) => [k, fieldText(values[k] ?? null)]))

const sameTexts = (a: Texts, b: Texts, keys: readonly string[]) =>
  keys.every((k) => (a[k] ?? '') === (b[k] ?? ''))

/** Syncs the project with a partner's exchange model. */
export function syncProject(
  data: ProjectData,
  incoming: DvhModel,
  partner: string,
  ctx: ChangeContext & { readonly at: string },
): SyncResult {
  const types = new Map(projectTypes(data).map((t) => [t.name, t]))
  const keysOf = (type: string) => types.get(type)?.fields.map((f) => f.key) ?? []
  const base = data.sync[partner]?.objects ?? {}
  const remotes = remoteObjects(data, incoming)
  // a one-per-project object the partner has but this project does not: its id comes from the base
  const singleBaseId = (type: string) =>
    Object.entries(base).find(([, b]) => b.type === type && types.get(type)?.single)?.[0]
  const remoteById = new Map<string, RemoteObject>()
  const newRemote: RemoteObject[] = []
  for (const r of remotes) {
    const id = r.id ?? (types.get(r.type)?.single ? (singleBaseId(r.type) ?? null) : null)
    if (id) remoteById.set(id, { ...r, id })
    else newRemote.push(r)
  }

  let next = data
  const conflicts: SyncConflict[] = []
  const agreed: SyncBase['objects'] = {}
  /** values the partner keeps for itself (conflicts), by object */
  const keepRemote = new Map<string, Texts>()
  /** objects the outgoing model leaves out / adds back */
  const omit = new Set<string>()
  const addBack: RemoteObject[] = []
  const stats = { pulled: 0, pushed: 0, created: 0, deleted: 0 }

  const syncable = (type: string) => types.has(type)
  const localIds = new Set(data.objects.filter((o) => syncable(o.type)).map((o) => o.id))
  const ids = new Set([...localIds, ...remoteById.keys(), ...Object.keys(base)])

  for (const id of ids) {
    const local = next.objects.find((o) => o.id === id && syncable(o.type))
    const remote = remoteById.get(id)
    const b = base[id]
    const type = local?.type ?? remote?.type ?? b?.type
    if (!type || !syncable(type)) continue
    const keys = keysOf(type)

    if (local && remote) {
      const localTexts = textsOf(local.values, keys)
      const pull: Record<string, Scalar> = {}
      const agreedValues: Texts = {}
      for (const key of keys) {
        const l = localTexts[key] ?? ''
        const r = remote.values[key]
        if (r === undefined) {
          // the partner does not carry this value: nothing to compare
          if (b?.values[key] !== undefined) agreedValues[key] = b.values[key]!
          continue
        }
        const bv = b?.values[key]
        if (l === r) {
          agreedValues[key] = l
        } else if (bv === undefined ? l === '' : l === bv) {
          pull[key] = r
          agreedValues[key] = r
          stats.pulled++
        } else if (bv === undefined ? r === '' : r === bv) {
          agreedValues[key] = l
          stats.pushed++
        } else {
          conflicts.push({
            objectId: id,
            type,
            key,
            kind: 'value',
            base: bv ?? null,
            local: l,
            remote: r,
          })
          keepRemote.set(id, { ...keepRemote.get(id), [key]: r })
          if (bv !== undefined) agreedValues[key] = bv
        }
      }
      if (Object.keys(pull).length) next = setValues(next, id, pull, ctx)
      agreed[id] = { type, values: agreedValues }
      continue
    }

    if (local && !remote) {
      if (!b) {
        stats.pushed++ // new here: the partner gets it
        continue
      }
      if (sameTexts(textsOf(local.values, keys), b.values, keys)) {
        next = deleteObject(next, id, ctx)
        stats.deleted++
      } else {
        conflicts.push({
          objectId: id,
          type,
          kind: 'deleted-there',
          base: null,
          local: JSON.stringify(textsOf(local.values, keys)),
          remote: null,
        })
        omit.add(id)
        agreed[id] = b
      }
      continue
    }

    if (!local && remote) {
      if (!b) {
        // a partner id this store cannot hold gets a new one, which goes back with the outgoing model
        const created = createObject(
          next,
          type,
          remote.values,
          ctx,
          /^[A-Za-z0-9_-]{1,64}$/.test(id) ? id : undefined,
        )
        next = created.data
        agreed[created.object.id] = { type, values: remote.values }
        stats.created++
        continue
      }
      if (sameTexts(remote.values, b.values, keys)) {
        // deleted here, untouched there: the partner loses it too
        stats.deleted++
      } else {
        conflicts.push({
          objectId: id,
          type,
          kind: 'deleted-here',
          base: null,
          local: null,
          remote: JSON.stringify(remote.values),
          remoteValues: remote.values,
        })
        addBack.push(remote)
        agreed[id] = b
      }
    }
  }

  // rows the partner added without an id: new objects here, ids go back with the outgoing model
  for (const r of newRemote) {
    const created = createObject(next, r.type, r.values, ctx)
    next = created.data
    agreed[created.object.id] = { type: r.type, values: r.values }
    stats.created++
  }

  next = {
    ...next,
    sync: {
      ...next.sync,
      [partner]: {
        at: ctx.at,
        objects: agreed,
        // a new sync replaces the old list: what is still in conflict was found again
        ...(conflicts.length ? { conflicts: structuredClone(conflicts) } : {}),
      },
    },
  }
  return { data: next, outgoing: outgoingModel(next, keepRemote, omit, addBack), conflicts, stats }
}

/** The merged project as the partner should store it, its conflicting values left as they are. */
function outgoingModel(
  data: ProjectData,
  keepRemote: ReadonlyMap<string, Texts>,
  omit: ReadonlySet<string>,
  addBack: readonly RemoteObject[],
): DvhModel {
  // the partner's own texts, typed again (a number stays a number in its workbook)
  const typed = (type: string, texts: Texts): Record<string, Scalar> => {
    const def = typeOf(data, type)
    return Object.fromEntries(
      Object.entries(texts).map(([key, value]) => {
        const field = def?.fields.find((f) => f.key === key)
        return [key, field ? coerceValue(field, value) : value]
      }),
    )
  }
  const objects = data.objects
    .filter((o) => !omit.has(o.id))
    .map((o) => {
      const kept = keepRemote.get(o.id)
      return kept ? { ...o, values: { ...o.values, ...typed(o.type, kept) } } : o
    })
  for (const r of addBack) {
    objects.push({ id: r.id!, type: r.type, rev: 0, values: typed(r.type, r.values), revs: {} })
  }
  return projectSourceModel({ ...data, objects })
}

/** The conflicts of a partner still waiting for the user. */
export function pendingConflicts(data: ProjectData, partner: string): readonly SyncConflict[] {
  return data.sync[partner]?.conflicts ?? []
}

const sameConflict = (a: SyncConflict, b: SyncConflict) =>
  a.objectId === b.objectId && a.kind === b.kind && (a.key ?? '') === (b.key ?? '')

/**
 * Settles every pending conflict of a partner the same way. An object whose
 * value conflict was settled is kept consistent: conflicts on objects that a
 * deletion conflict already removed are skipped.
 */
export function resolveAllSyncConflicts(
  data: ProjectData,
  partner: string,
  keep: 'mine' | 'theirs',
  ctx: ChangeContext,
): ProjectData {
  let next = data
  for (const conflict of pendingConflicts(data, partner)) {
    if (!pendingConflicts(next, partner).some((c) => sameConflict(c, conflict))) continue
    next = resolveSyncConflict(next, partner, conflict, keep, ctx)
  }
  return next
}

/** Settles one conflict: `mine` is pushed at the next sync, `theirs` is taken now. */
export function resolveSyncConflict(
  data: ProjectData,
  partner: string,
  conflict: SyncConflict,
  keep: 'mine' | 'theirs',
  ctx: ChangeContext,
): ProjectData {
  const sync = data.sync[partner]
  if (!sync) return data
  const objects = { ...sync.objects }
  let next = data
  // a conflict that is no longer pending (settled, or replaced by a newer sync) changes nothing
  if (sync.conflicts && !sync.conflicts.some((c) => sameConflict(c, conflict))) return data
  const remaining = (sync.conflicts ?? []).filter((c) => !sameConflict(c, conflict))
  if (conflict.kind === 'value' && conflict.key) {
    const key = conflict.key
    const remote = conflict.remote ?? ''
    const entry = objects[conflict.objectId] ?? { type: conflict.type, values: {} }
    // the partner's value becomes the base: unchanged there, so "mine" wins next time
    objects[conflict.objectId] = { type: entry.type, values: { ...entry.values, [key]: remote } }
    if (keep === 'theirs') {
      next = next.objects.some((o) => o.id === conflict.objectId)
        ? setValues(next, conflict.objectId, { [key]: remote }, ctx)
        : // deleted here meanwhile: theirs brings it back with the partner's value
          createObject(
            next,
            conflict.type,
            { ...entry.values, [key]: remote },
            ctx,
            conflict.objectId,
          ).data
    }
  } else if (conflict.kind === 'deleted-there') {
    delete objects[conflict.objectId]
    if (keep === 'theirs') next = deleteObject(next, conflict.objectId, ctx)
  } else if (conflict.kind === 'deleted-here') {
    const values = conflict.remoteValues ?? {}
    if (keep === 'theirs') {
      next = next.objects.some((o) => o.id === conflict.objectId)
        ? setValues(next, conflict.objectId, values, ctx)
        : createObject(next, conflict.type, values, ctx, conflict.objectId).data
    }
    // either way the base is the partner's object: kept here, or deleted there next time
    objects[conflict.objectId] = { type: conflict.type, values: { ...values } }
  }
  const { conflicts: _old, ...rest } = sync
  return {
    ...next,
    sync: {
      ...next.sync,
      [partner]: { ...rest, objects, ...(remaining.length ? { conflicts: remaining } : {}) },
    },
  }
}

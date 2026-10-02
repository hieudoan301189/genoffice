/**
 * Documents link straight to project objects (P10). The project shows itself
 * to a document as a virtual link source `dvh-project://<projectId>` whose
 * Smart Data model is derived from the objects:
 *
 * - each `single` type (one object per project) gives fields `Type.key`;
 * - each list type gives a collection `Type` with an `ID` column first, one
 *   row per object.
 *
 * Ids are stable (`pf_<objectId>_<key>`, `pc_<Type>`), so the P3 machinery —
 * three-way pull, revisions, write-back with conflict checks — works on
 * project data unchanged.
 */

import {
  checkFieldWrite,
  emptyModel,
  fieldText,
  type DvhCollection,
  type DvhField,
  type DvhModel,
  type FieldType,
  type FieldWrite,
  type FieldWriteResult,
} from '@genoffice/dvh-model'
import { projectTypes, setValues, type ChangeContext, type ProjectData } from './data'
import type { PackField } from './pack'

export const PROJECT_URI_PREFIX = 'dvh-project://'

export const projectUri = (projectId: string) =>
  `${PROJECT_URI_PREFIX}${encodeURIComponent(projectId)}`

/** The project id of a `dvh-project://` path, or null for any other path. */
export function projectIdOfUri(path: string): string | null {
  if (!path.startsWith(PROJECT_URI_PREFIX)) return null
  try {
    const id = decodeURIComponent(path.slice(PROJECT_URI_PREFIX.length))
    return id && !/[\\/]/.test(id) ? id : null
  } catch {
    return null
  }
}

export const projectFieldId = (objectId: string, key: string) => `pf_${objectId}_${key}`
export const projectCollectionId = (type: string) => `pc_${type}`
export const projectColumnId = (type: string, key: string) => `pc_${type}_${key}`

const fieldTypeOf = (field: PackField): FieldType => (field.type === 'ref' ? 'text' : field.type)

/** The project as a Smart Data model (what a linked document reads). */
export function projectSourceModel(data: ProjectData): DvhModel {
  const model = emptyModel(data.docId)
  const fields: DvhField[] = []
  const collections: DvhCollection[] = []
  for (const type of projectTypes(data)) {
    const objects = data.objects.filter((o) => o.type === type.name)
    if (type.single) {
      const object = objects[0]
      if (!object) continue
      for (const field of type.fields) {
        fields.push({
          id: projectFieldId(object.id, field.key),
          name: `${type.name}.${field.key}`,
          type: fieldTypeOf(field),
          value: object.values[field.key] ?? null,
          ...(field.enumValues ? { enumValues: field.enumValues } : {}),
          access: 'readwrite',
          rev: object.revs[field.key] ?? 0,
        })
      }
      continue
    }
    collections.push({
      id: projectCollectionId(type.name),
      name: type.name,
      columns: [
        { id: projectColumnId(type.name, 'id'), key: 'id', title: 'ID', type: 'text' },
        ...type.fields.map((field) => ({
          id: projectColumnId(type.name, field.key),
          key: field.key,
          title: field.label ?? field.key,
          type: fieldTypeOf(field),
        })),
      ],
      rows: objects.map((o) => [o.id, ...type.fields.map((f) => o.values[f.key] ?? null)]),
    })
  }
  return { ...model, fields, collections }
}

/** The object and key a project field id stands for. */
export function resolveProjectField(
  data: ProjectData,
  fieldId: string,
): { objectId: string; key: string } | null {
  for (const object of data.objects) {
    const prefix = `pf_${object.id}_`
    if (fieldId.startsWith(prefix))
      return { objectId: object.id, key: fieldId.slice(prefix.length) }
  }
  return null
}

/**
 * Field write-back into the project (the P3 protocol): each write is checked
 * against the value's revision and text as the document last saw them; a
 * value changed in the project since then is a conflict, never overwritten.
 */
export function writeProjectFields(
  data: ProjectData,
  writes: readonly FieldWrite[],
  ctx: ChangeContext,
): { data: ProjectData; results: FieldWriteResult[] } {
  let next = data
  const results: FieldWriteResult[] = []
  for (const write of writes) {
    const target = resolveProjectField(next, write.fieldId)
    const object = target && next.objects.find((o) => o.id === target.objectId)
    if (!target || !object) {
      results.push({ fieldId: write.fieldId, status: 'missing' })
      continue
    }
    const current = {
      rev: object.revs[target.key] ?? 0,
      text: fieldText(object.values[target.key] ?? null),
    }
    if (checkFieldWrite(write, current) === 'conflict') {
      results.push({ fieldId: write.fieldId, status: 'conflict', current })
      continue
    }
    next = setValues(next, object.id, { [target.key]: write.value }, ctx)
    const after = next.objects.find((o) => o.id === object.id)!
    results.push({
      fieldId: write.fieldId,
      status: 'written',
      current: {
        rev: after.revs[target.key] ?? 0,
        text: fieldText(after.values[target.key] ?? null),
      },
    })
  }
  return { data: next, results }
}

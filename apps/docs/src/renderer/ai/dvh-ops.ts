import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { fieldText, newDvhId, scalarSchema, type DvhField, type Scalar } from '@genoffice/dvh-model'
import type { Op, OpContext, OpDef, OpResult, RunEnv, Target, TopBlock } from './ops'
import { posAfterText } from './after-text'
import {
  activeDvhDocs,
  ensureModel,
  fieldOccurrences,
  findField,
  noteFieldInserted,
  replaceFieldTexts,
  setFieldValue,
  smartFieldNode,
} from '../dvh-smart-data'

/**
 * DVH Smart Data ops (P1): insert a field occurrence and set a field's value,
 * so the agent, the CLI and MCP get them through apply_ops. Text changes go
 * into the batch transaction; the model is written in commit(), after the
 * batch applied, so a rejected batch leaves the model untouched.
 */

export interface DvhOpHelpers {
  validateShape(op: Op, def: OpDef, where: string): string | null
  matchTarget(doc: PmNode, target: Target, sel: RunEnv['sel']): TopBlock[]
  markChanged(tr: Transaction, pos: number, ctx: RunEnv['ctx']): void
}

const TEXT_BLOCKS = new Set(['docParagraph', 'docHeading', 'docListItem'])
/** fields created by an insertSmartField op, carried from apply() to commit() */
const created = new WeakMap<Op, DvhField>()
/** fields created earlier in the same batch, so later ops can use them by name */
const createdInBatch = new WeakMap<Transaction, DvhField[]>()

function batchField(tr: Transaction, ref: string): DvhField | undefined {
  return (
    findField(activeDvhDocs(), ref) ??
    createdInBatch.get(tr)?.find((f) => f.id === ref || f.name === ref)
  )
}

const unknownField = (name: string, ref: string) =>
  new Error(`${name}: unknown field "${ref}"; known fields: ${knownFields()}`)

const sourceOf = (ctx: OpContext) => (ctx.source === 'ui' ? 'ui' : 'ai')

function knownFields(): string {
  const names = activeDvhDocs()?.model?.fields.map((f) => f.name) ?? []
  return names.length > 0 ? names.join(', ') : 'none (link a workbook or give value to create one)'
}

function validateValue(op: Op, where: string): string | null {
  return op.value === undefined || scalarSchema.safeParse(op.value).success
    ? null
    : `${where}: value must be a string, number, boolean or null`
}

export function dvhOpDefs(h: DvhOpHelpers): OpDef[] {
  const insertSmartField: OpDef = {
    name: 'insertSmartField',
    signature:
      '{ op: "insertSmartField", target, field, value?, position?: "start"|"end", afterText?: string }  // insert a Smart Data field into exactly one paragraph, at its end unless position or afterText says otherwise: a content control bound to the document\'s data, so Word shows the same value and setSmartField / "Update from source" refresh every occurrence. field is a field name such as "Project.Name" (or its id); a name not in the document yet needs value and creates a text field',
    keys: ['field', 'value', 'position', 'afterText'],
    target: 'required',
    validate(op, where) {
      const shape = h.validateShape(op, insertSmartField, where)
      if (shape) return shape
      if (typeof op.field !== 'string' || op.field.trim() === '')
        return `${where}: field must be a field name`
      const value = validateValue(op, where)
      if (value) return value
      if (op.position !== undefined && op.position !== 'start' && op.position !== 'end')
        return `${where}: position must be "start" or "end"`
      if (op.afterText !== undefined && (typeof op.afterText !== 'string' || op.afterText === ''))
        return `${where}: afterText must be a non-empty string`
      if (op.afterText !== undefined && op.position !== undefined)
        return `${where}: give either position or afterText, not both`
      return null
    },
    apply(op, env) {
      const dvh = activeDvhDocs()
      if (!dvh) throw new Error('insertSmartField: no document is open')
      const ready = ensureModel(dvh)
      const ref = String(op.field).trim()
      let field = batchField(env.tr, ref)
      if (!field) {
        if (op.value === undefined) throw unknownField('insertSmartField', ref)
        field = {
          id: newDvhId('f'),
          name: ref,
          type:
            typeof op.value === 'number'
              ? 'number'
              : typeof op.value === 'boolean'
                ? 'boolean'
                : 'text',
          value: op.value as Scalar,
          access: 'readwrite',
        }
        created.set(op, field)
        createdInBatch.set(env.tr, [...(createdInBatch.get(env.tr) ?? []), field])
      }
      const blocks = h.matchTarget(env.tr.doc, op.target as Target, env.sel)
      if (blocks.length !== 1) {
        throw new Error(
          `insertSmartField: target must match exactly one paragraph (matched ${blocks.length}); use blockIndexes`,
        )
      }
      const b = blocks[0]!
      if (!TEXT_BLOCKS.has(b.node.type.name))
        throw new Error(
          `insertSmartField: block ${b.index} is not a paragraph, heading or list item`,
        )
      const contentStart = b.pos + 1
      let pos = op.position === 'start' ? contentStart : b.pos + b.node.nodeSize - 1
      if (typeof op.afterText === 'string') {
        const where = posAfterText(b.node, contentStart, op.afterText, `block ${b.index}`)
        if ('error' in where) throw new Error(`insertSmartField: ${where.error}`)
        pos = where.pos
      }
      env.tr.insert(pos, smartFieldNode(env.schema, env.tr.doc, ready, field))
      h.markChanged(env.tr, b.pos, env.ctx)
      return {
        op: 'insertSmartField',
        matched: 1,
        changed: 1,
        skippedProtected: 0,
        detail: `${field.name} = "${fieldText(field.value)}"`,
      }
    },
    commit(op, ctx) {
      const dvh = activeDvhDocs()
      const field = created.get(op) ?? findField(dvh, String(op.field).trim())
      if (dvh && field) noteFieldInserted(dvh, field, sourceOf(ctx))
    },
  }

  const setSmartField: OpDef = {
    name: 'setSmartField',
    signature:
      '{ op: "setSmartField", field, value }  // set a Smart Data field\'s value (field name such as "Project.Name", or its id); every occurrence in the document shows the new value and Word sees it through the data binding',
    keys: ['field', 'value'],
    target: 'none',
    validate(op, where) {
      const shape = h.validateShape(op, setSmartField, where)
      if (shape) return shape
      if (typeof op.field !== 'string' || op.field.trim() === '')
        return `${where}: field must be a field name`
      if (op.value === undefined) return `${where}: value is required`
      return validateValue(op, where)
    },
    apply(op, env): OpResult {
      const ref = String(op.field).trim()
      const field = batchField(env.tr, ref)
      if (!field) throw unknownField('setSmartField', ref)
      const occurrences = fieldOccurrences(env.tr.doc).filter((o) => o.fieldId === field.id)
      const changed = replaceFieldTexts(
        env.tr,
        new Map([[field.id, fieldText(op.value as Scalar)]]),
      )
      return {
        op: 'setSmartField',
        matched: occurrences.length,
        changed,
        skippedProtected: 0,
        detail: `${field.name} = "${fieldText(op.value as Scalar)}" (${occurrences.length} in the document)`,
      }
    },
    commit(op, ctx) {
      const dvh = activeDvhDocs()
      const field = findField(dvh, String(op.field).trim())
      if (dvh && field)
        setFieldValue(dvh, field, op.value as Scalar, 'Data.SetField', sourceOf(ctx))
    },
  }

  return [insertSmartField, setSmartField]
}

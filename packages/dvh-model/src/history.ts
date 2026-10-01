/**
 * Change sets (ADR D7) and the embedded history part (ADR D9):
 * `<dvh:history xmlns:dvh="urn:dvh-office:history:1">` holding one JSON change
 * set per line in CDATA (the zip already compresses it), plus the model hash at
 * the last DVH save so an edit made in Word/Excel can be detected on open.
 */

import { z } from 'zod'

import { decodeXml, escapeXml, HISTORY_NS } from './custom-xml'
import type { DvhModel } from './types'

export const changeSetSchema = z
  .object({
    id: z.string().min(1),
    txId: z.string().min(1),
    docId: z.string().min(1),
    at: z.string(),
    source: z.enum(['ui', 'ai', 'workflow', 'script', 'link', 'restore', 'external']),
    actor: z
      .object({ user: z.string().optional(), device: z.string().optional() })
      .strict()
      .optional(),
    action: z.string().min(1),
    changes: z.array(
      z
        .object({
          objectId: z.string().min(1),
          path: z.string(),
          before: z.unknown(),
          after: z.unknown(),
        })
        .strict(),
    ),
  })
  .strict()
export type ChangeSet = z.infer<typeof changeSetSchema>

export interface HistoryPart {
  readonly docId: string
  /** modelHash() of the model at the last DVH save */
  readonly modelHash?: string
  readonly changes: readonly ChangeSet[]
}

export function serializeHistoryXml(history: HistoryPart): string {
  const lines = history.changes.map((c) => JSON.stringify(c)).join('\n')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<dvh:history xmlns:dvh="${HISTORY_NS}" docId="${escapeXml(history.docId)}"` +
    (history.modelHash ? ` modelHash="${history.modelHash}"` : '') +
    ` count="${history.changes.length}">` +
    `<dvh:changes><![CDATA[${lines.replace(/\]\]>/g, ']]]]><![CDATA[>')}]]></dvh:changes>` +
    '</dvh:history>'
  )
}

export function parseHistoryXml(xml: string): HistoryPart {
  const root = /<(?:\w+:)?history\b[^>]*>/.exec(xml)
  if (!root || !xml.includes(HISTORY_NS)) throw new Error('not a DVH history part')
  const attr = (name: string) => new RegExp(`\\s${name}="([^"]*)"`).exec(root[0])?.[1]
  const body = /<(?:\w+:)?changes\b[^>]*>([\s\S]*?)<\/(?:\w+:)?changes>/.exec(xml)?.[1] ?? ''
  const changes = decodeXml(body)
    .split('\n')
    .filter((line) => line.trim() !== '')
    .map((line) => changeSetSchema.parse(JSON.parse(line)))
  const modelHash = attr('modelHash')
  return { docId: decodeXml(attr('docId') ?? ''), ...(modelHash ? { modelHash } : {}), changes }
}

/** Order-independent JSON for hashing. */
function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(',')}]`
  if (value && typeof value === 'object') {
    return `{${Object.keys(value as Record<string, unknown>)
      .sort()
      .map((k) => `${JSON.stringify(k)}:${canonical((value as Record<string, unknown>)[k])}`)
      .join(',')}}`
  }
  return JSON.stringify(value) ?? 'null'
}

/** 64-bit FNV-1a of the canonical model, hex; synchronous so renderers can call it. */
export function modelHash(model: DvhModel): string {
  let hash = 0xcbf29ce484222325n
  const prime = 0x100000001b3n
  for (const byte of new TextEncoder().encode(canonical(model))) {
    hash ^= BigInt(byte)
    hash = (hash * prime) & 0xffffffffffffffffn
  }
  return hash.toString(16).padStart(16, '0')
}

/** Field-level differences between two models, as change-set entries. */
export function diffModels(before: DvhModel, after: DvhModel): ChangeSet['changes'] {
  const changes: ChangeSet['changes'] = []
  const old = new Map(before.fields.map((f) => [f.id, f]))
  for (const field of after.fields) {
    const prev = old.get(field.id)
    if (!prev) changes.push({ objectId: field.id, path: '', before: null, after: field })
    else {
      if (prev.value !== field.value)
        changes.push({ objectId: field.id, path: 'value', before: prev.value, after: field.value })
      if (prev.name !== field.name)
        changes.push({ objectId: field.id, path: 'name', before: prev.name, after: field.name })
    }
    old.delete(field.id)
  }
  for (const removed of old.values())
    changes.push({ objectId: removed.id, path: '', before: removed, after: null })
  return changes
}

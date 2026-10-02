/**
 * The action bridge (P4): how a process that does not hold a document (the
 * shell's MCP tools) reaches the catalog of one that does (a Docs or Sheets
 * renderer, the Slides session in main). One request shape for every app:
 *
 * - `catalog`: the actions, their effects and JSON Schema inputs, plus the
 *   catalog fingerprint and the document's docId;
 * - `run`: one action, dry or real, with the permissions the caller was granted;
 * - `checkpoint` / `restore` / `release`: the saga hooks. Checkpoints stay in
 *   the document's process; only an opaque token crosses the bridge.
 *
 * Requests are validated here, on the receiving side.
 */

import { z } from 'zod'
import type { ActionEffect, ActionRegistry, CatalogEntry, RunResult } from './index'
import type { SagaParticipant } from './saga'

const effectSchema = z.enum(['read', 'write', 'bulk', 'destructive', 'external'])

export const dvhBridgeRequestSchema = z.discriminatedUnion('verb', [
  z.object({ verb: z.literal('catalog') }).strict(),
  z
    .object({
      verb: z.literal('run'),
      name: z.string().min(1).max(128),
      input: z.unknown(),
      caller: z.enum(['ui', 'ai', 'workflow', 'script', 'extension']),
      dryRun: z.boolean().optional(),
      txId: z.string().min(1).max(64).optional(),
      fingerprint: z.string().max(64).optional(),
      /** effects the caller may perform for this run */
      allow: z.array(effectSchema).max(5),
      /** the user already confirmed this run (the caller showed the preview) */
      confirmed: z.boolean(),
      bulkThreshold: z.number().int().min(0).optional(),
    })
    .strict(),
  z.object({ verb: z.literal('checkpoint') }).strict(),
  z.object({ verb: z.literal('restore'), token: z.string().min(1).max(64) }).strict(),
  z.object({ verb: z.literal('release'), token: z.string().min(1).max(64) }).strict(),
])
export type DvhBridgeRequest = z.infer<typeof dvhBridgeRequestSchema>

export interface DvhCatalogReply {
  readonly docId: string
  readonly fingerprint: string
  readonly actions: readonly CatalogEntry[]
}

export interface DvhBridgeTarget {
  readonly docId: string
  readonly registry: ActionRegistry
  readonly participant: SagaParticipant
}

/** checkpoints kept per process; a saga releases (or restores) its own */
const MAX_CHECKPOINTS = 32

/**
 * The receiving end: validates a request and answers it against the document
 * `target()` returns at that moment (null: no document open).
 */
export function createBridgeEndpoint(
  target: () => DvhBridgeTarget | null,
): (request: unknown) => Promise<unknown> {
  const checkpoints = new Map<string, unknown>()
  let seq = 0
  return async (raw) => {
    const parsed = dvhBridgeRequestSchema.safeParse(raw)
    if (!parsed.success) throw new Error(`invalid DVH bridge request: ${parsed.error.message}`)
    const request = parsed.data
    const current = target()
    if (!current) throw new Error('no document is open')
    switch (request.verb) {
      case 'catalog':
        return {
          docId: current.docId,
          fingerprint: current.registry.fingerprint(),
          actions: current.registry.catalog(),
        } satisfies DvhCatalogReply
      case 'run': {
        const result: RunResult = await current.registry.run(request.name, request.input, {
          caller: request.caller,
          docId: current.docId,
          permissions: new Set<ActionEffect>(request.allow),
          ...(request.dryRun ? { dryRun: true } : {}),
          ...(request.txId ? { txId: request.txId } : {}),
          ...(request.fingerprint ? { fingerprint: request.fingerprint } : {}),
          ...(request.bulkThreshold !== undefined
            ? { policy: { bulkThreshold: request.bulkThreshold } }
            : {}),
          confirm: () => request.confirmed,
        })
        return result
      }
      case 'checkpoint': {
        if (checkpoints.size >= MAX_CHECKPOINTS) {
          checkpoints.delete(checkpoints.keys().next().value!)
        }
        const token = `cp_${Date.now().toString(36)}_${(seq++).toString(36)}`
        checkpoints.set(token, await current.participant.checkpoint())
        return { token }
      }
      case 'restore': {
        if (!checkpoints.has(request.token)) throw new Error('unknown or expired checkpoint')
        await current.participant.restore(checkpoints.get(request.token))
        checkpoints.delete(request.token)
        return { restored: true }
      }
      case 'release':
        checkpoints.delete(request.token)
        return { released: true }
    }
  }
}

/**
 * The calling end: a saga participant whose every call crosses the bridge
 * (`send` reaches one document's endpoint).
 */
export function remoteParticipant(options: {
  docId: string
  label?: string
  send(request: DvhBridgeRequest): Promise<unknown>
  /** what the caller may do in this saga, and whether the user confirmed it */
  grant: {
    allow: readonly ActionEffect[]
    confirmed: boolean
    bulkThreshold?: number
    fingerprint?: string
  }
}): SagaParticipant & { release(token: unknown): Promise<void> } {
  const { send, grant } = options
  return {
    docId: options.docId,
    ...(options.label ? { label: options.label } : {}),
    run: async (name, input, run) =>
      (await send({
        verb: 'run',
        name,
        input,
        caller: run.caller,
        ...(run.dryRun ? { dryRun: true } : {}),
        txId: run.txId,
        ...(grant.fingerprint ? { fingerprint: grant.fingerprint } : {}),
        allow: [...grant.allow],
        confirmed: grant.confirmed,
        ...(grant.bulkThreshold !== undefined ? { bulkThreshold: grant.bulkThreshold } : {}),
      })) as RunResult,
    checkpoint: async () => ((await send({ verb: 'checkpoint' })) as { token: string }).token,
    restore: async (token) => {
      await send({ verb: 'restore', token: String(token) })
    },
    release: async (token) => {
      await send({ verb: 'release', token: String(token) })
    },
  }
}

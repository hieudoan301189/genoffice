/**
 * P4 bridge: a saga driven from another process (the shell) through the
 * request protocol behaves like a local one — refusals included.
 */
import { describe, expect, it } from 'vitest'
import { z } from 'zod'
import {
  ActionRegistry,
  createBridgeEndpoint,
  localParticipant,
  registerOpFamily,
  remoteParticipant,
  runSaga,
  type DvhBridgeRequest,
} from '../src/index'

function remoteDoc(docId: string) {
  let text = 'v0'
  const registry = new ActionRegistry()
  registerOpFamily(registry, {
    group: 'Document',
    entries: [
      { op: 'setText', input: z.object({ text: z.string() }).strict() },
      { op: 'deleteAll', input: z.object({}).strict() },
    ],
    run: async (op, dryRun) => {
      if (op.op === 'setText' && op.text === 'boom') throw new Error('rejected')
      if (!dryRun) text = op.op === 'deleteAll' ? '' : String(op.text)
      return { summary: [String(op.op)], objects: 1 }
    },
  })
  const participant = localParticipant({
    docId,
    registry,
    checkpoint: () => text,
    restore: (cp) => {
      text = String(cp)
    },
  })
  const endpoint = createBridgeEndpoint(() => ({ docId, registry, participant }))
  // what crosses IPC is plain JSON
  const send = async (request: DvhBridgeRequest) =>
    JSON.parse(JSON.stringify(await endpoint(JSON.parse(JSON.stringify(request))))) as unknown
  return { send, text: () => text, registry }
}

describe('bridge', () => {
  it('lists the catalog with its fingerprint and refuses malformed requests', async () => {
    const doc = remoteDoc('doc_a')
    const catalog = (await doc.send({ verb: 'catalog' })) as {
      docId: string
      fingerprint: string
      actions: { name: string }[]
    }
    expect(catalog.docId).toBe('doc_a')
    expect(catalog.fingerprint).toBe(doc.registry.fingerprint())
    expect(catalog.actions.map((a) => a.name)).toEqual(['Document.SetText', 'Document.DeleteAll'])
    await expect(doc.send({ verb: 'nope' } as never)).rejects.toThrow(/invalid DVH bridge request/)
  })

  it('a remote saga rolls back across the bridge; grants decide destructive runs', async () => {
    const a = remoteDoc('doc_a')
    const b = remoteDoc('doc_b')
    const grant = { allow: ['read', 'write'] as const, confirmed: false }
    const participants = new Map([
      [
        'doc_a',
        remoteParticipant({
          docId: 'doc_a',
          send: a.send,
          grant: { ...grant, allow: [...grant.allow] },
        }),
      ],
      [
        'doc_b',
        remoteParticipant({
          docId: 'doc_b',
          send: b.send,
          grant: { ...grant, allow: [...grant.allow] },
        }),
      ],
    ])
    const options = { caller: 'ai' as const, permissions: new Set<never>() }
    const failed = await runSaga(
      [
        { docId: 'doc_a', action: 'Document.SetText', input: { text: 'a1' } },
        { docId: 'doc_b', action: 'Document.SetText', input: { text: 'boom' } },
      ],
      participants,
      options,
    )
    expect(failed.ok).toBe(false)
    expect(a.text()).toBe('v0')
    // without the grant, a destructive step is refused before it runs
    const refused = await runSaga(
      [{ docId: 'doc_a', action: 'Document.DeleteAll', input: {} }],
      participants,
      options,
    )
    expect(refused.ok).toBe(false)
    expect(refused.error).toMatch(/may not run destructive/)
    expect(a.text()).toBe('v0')
    const granted = new Map([
      [
        'doc_a',
        remoteParticipant({
          docId: 'doc_a',
          send: a.send,
          grant: { allow: ['read', 'write', 'destructive'], confirmed: true },
        }),
      ],
    ])
    const ok = await runSaga(
      [{ docId: 'doc_a', action: 'Document.DeleteAll', input: {} }],
      granted,
      options,
    )
    expect(ok.ok).toBe(true)
    expect(a.text()).toBe('')
  })
})

/**
 * P4: DVH actions over MCP. A plan spanning two open documents runs as one
 * saga (a failing step rolls both back); actions outside a document's catalog
 * are refused before anything runs; destructive steps and plans over the
 * threshold need the user's confirmation, and a declined plan changes nothing.
 */
import { describe, expect, it } from 'vitest'
import { z } from 'zod'
import {
  ActionRegistry,
  createBridgeEndpoint,
  localParticipant,
  registerOpFamily,
  type DvhBridgeRequest,
} from '@genoffice/dvh-actions'
import { createDvhTools, type DvhActionsControl } from '../../src/main/mcp/tools/dvh-tools'

function document(docId: string) {
  let lines = ['a']
  const registry = new ActionRegistry()
  registerOpFamily(registry, {
    group: 'Document',
    entries: [
      {
        op: 'append',
        input: z.object({ text: z.string(), count: z.number().int().optional() }).strict(),
      },
      { op: 'deleteAll', input: z.object({}).strict() },
    ],
    run: async (op, dryRun) => {
      if (op.text === 'boom' && !dryRun) throw new Error('write failed')
      const count = Number(op.count ?? 1)
      if (!dryRun) lines = op.op === 'deleteAll' ? [] : [...lines, String(op.text)]
      return { summary: [String(op.op)], objects: op.op === 'deleteAll' ? lines.length : count }
    },
  })
  const participant = localParticipant({
    docId,
    registry,
    checkpoint: () => [...lines],
    restore: (cp) => {
      lines = [...(cp as string[])]
    },
  })
  return {
    endpoint: createBridgeEndpoint(() => ({ docId, registry, participant })),
    lines: () => lines,
    registry,
  }
}

function setup(answer: boolean) {
  const docs = new Map([
    ['A.docx', { wcId: 1, doc: document('doc_a') }],
    ['B.xlsx', { wcId: 2, doc: document('doc_b') }],
  ])
  const asked: string[][] = []
  const control: DvhActionsControl = {
    resolve: async (name) => {
      const entry = docs.get(name)
      if (!entry) throw new Error(`no open document matches "${name}"`)
      return { wcId: entry.wcId, family: name.endsWith('.xlsx') ? 'xlsx' : 'docx', title: name }
    },
    send: async (wcId, _family, request: DvhBridgeRequest) => {
      const entry = [...docs.values()].find((d) => d.wcId === wcId)!
      return JSON.parse(
        JSON.stringify(await entry.doc.endpoint(JSON.parse(JSON.stringify(request)))),
      )
    },
    confirm: async (_title, lines) => {
      asked.push([...lines])
      return answer
    },
    bulkThreshold: () => 10,
  }
  const tools = Object.fromEntries(createDvhTools(control).map((t) => [t.name, t]))
  return { tools, asked, a: docs.get('A.docx')!.doc, b: docs.get('B.xlsx')!.doc }
}

describe('dvh MCP tools', () => {
  it('lists a document catalog with its fingerprint', async () => {
    const { tools, a } = setup(true)
    const listed = (await tools.dvh_list_actions!.handler({ document: 'A.docx' })) as {
      docId: string
      fingerprint: string
      actions: { name: string; effect: string }[]
    }
    expect(listed.docId).toBe('doc_a')
    expect(listed.fingerprint).toBe(a.registry.fingerprint())
    expect(listed.actions.map((x) => `${x.name}:${x.effect}`)).toEqual([
      'Document.Append:write',
      'Document.DeleteAll:destructive',
    ])
  })

  it('runs a two-document plan as one saga and rolls both back on failure', async () => {
    const { tools, a, b, asked } = setup(true)
    const ok = (await tools.dvh_execute!.handler({
      steps: [
        { document: 'A.docx', action: 'Document.Append', input: { text: 'x' } },
        { document: 'B.xlsx', action: 'Document.Append', input: { text: 'y' } },
      ],
    })) as { ok: boolean; changeSets: number }
    expect(ok).toMatchObject({ ok: true, changeSets: 2 })
    expect(a.lines()).toEqual(['a', 'x'])
    expect(b.lines()).toEqual(['a', 'y'])
    expect(asked).toEqual([])

    const failed = (await tools.dvh_execute!.handler({
      steps: [
        { document: 'A.docx', action: 'Document.Append', input: { text: 'x2' } },
        { document: 'B.xlsx', action: 'Document.Append', input: { text: 'boom' } },
      ],
    })) as { ok: boolean; failedAt: number; rolledBack: string[] }
    expect(failed).toMatchObject({ ok: false, failedAt: 2, rolledBack: ['B.xlsx', 'A.docx'] })
    expect(a.lines()).toEqual(['a', 'x'])
    expect(b.lines()).toEqual(['a', 'y'])
  })

  it('refuses actions outside the catalog and stale fingerprints before running anything', async () => {
    const { tools, a } = setup(true)
    await expect(
      tools.dvh_execute!.handler({
        steps: [
          { document: 'A.docx', action: 'Document.Append', input: { text: 'x' } },
          { document: 'A.docx', action: 'File.DeleteDisk', input: {} },
        ],
      }),
    ).rejects.toThrow(/has no action File.DeleteDisk/)
    await expect(
      tools.dvh_execute!.handler({
        steps: [{ document: 'A.docx', action: 'Document.Append', input: { text: 'x' } }],
        fingerprints: { doc_a: '000000000000' },
      }),
    ).rejects.toThrow(/catalog of "A.docx" changed/)
    expect(a.lines()).toEqual(['a'])
  })

  it('asks before destructive or bulk plans; a declined plan changes nothing', async () => {
    const declined = setup(false)
    await expect(
      declined.tools.dvh_execute!.handler({
        steps: [{ document: 'A.docx', action: 'Document.DeleteAll', input: {} }],
      }),
    ).rejects.toThrow(/declined/)
    expect(declined.a.lines()).toEqual(['a'])
    expect(declined.asked[0]![0]).toContain('Document.DeleteAll')
    // over the threshold (10 objects), even for a write action
    await expect(
      declined.tools.dvh_execute!.handler({
        steps: [
          { document: 'A.docx', action: 'Document.Append', input: { text: 'x', count: 200 } },
        ],
      }),
    ).rejects.toThrow(/declined/)

    const accepted = setup(true)
    const preview = (await accepted.tools.dvh_preview!.handler({
      steps: [{ document: 'A.docx', action: 'Document.DeleteAll', input: {} }],
    })) as { needsConfirmation: boolean }
    expect(preview.needsConfirmation).toBe(true)
    const run = (await accepted.tools.dvh_execute!.handler({
      steps: [{ document: 'A.docx', action: 'Document.DeleteAll', input: {} }],
    })) as { ok: boolean }
    expect(run.ok).toBe(true)
    expect(accepted.a.lines()).toEqual([])
  })
})

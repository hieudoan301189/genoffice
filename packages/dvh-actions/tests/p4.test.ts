/**
 * P4 acceptance at the API level: a saga that fails halfway puts every
 * touched document back to its checkpoint; actions outside the catalog, or
 * beyond the caller's permissions, are refused; a stale plan is refused; the
 * confirmation policy follows the bulk threshold.
 */
import { describe, expect, it } from 'vitest'
import { z } from 'zod'
import {
  ActionError,
  ActionRegistry,
  defaultPermissions,
  effectOfOp,
  localParticipant,
  needsConfirmation,
  pascalCase,
  previewSaga,
  registerOpFamily,
  runSaga,
  type SagaParticipant,
} from '../src/index'

/** a toy document: a list of lines, one registry, snapshot checkpoints */
function toyDocument(docId: string) {
  let lines: string[] = ['a', 'b']
  const registry = new ActionRegistry()
  registerOpFamily(registry, {
    group: 'Document',
    entries: [
      { op: 'appendLine', input: z.object({ text: z.string() }).strict() },
      { op: 'deleteLine', input: z.object({ index: z.number().int() }).strict() },
      { op: 'failHere', input: z.object({}).strict() },
    ],
    run: async (op, dryRun) => {
      if (op.op === 'failHere') throw new Error('disk full')
      if (op.op === 'appendLine') {
        if (!dryRun) lines = [...lines, String(op.text)]
        return { summary: [`append "${String(op.text)}"`], objects: 1 }
      }
      const index = Number(op.index)
      if (index >= lines.length) throw new Error(`no line ${index}`)
      if (!dryRun) lines = lines.filter((_, i) => i !== index)
      return { summary: [`delete line ${index}`], objects: 1 }
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
  return { registry, participant, lines: () => lines }
}

const run = {
  caller: 'workflow' as const,
  permissions: defaultPermissions('ui'),
  confirm: () => true,
  now: () => '2026-10-02T00:00:00.000Z',
}

describe('adapters', () => {
  it('names and effects follow the op', () => {
    expect(pascalCase('setFont')).toBe('SetFont')
    expect(pascalCase('insert_rows')).toBe('InsertRows')
    expect(effectOfOp('deleteSheet')).toBe('destructive')
    expect(effectOfOp('clearRange')).toBe('destructive')
    expect(effectOfOp('setCells')).toBe('write')
    const doc = toyDocument('doc_a')
    expect(doc.registry.catalog().map((e) => `${e.name}:${e.effect}`)).toEqual([
      'Document.AppendLine:write',
      'Document.DeleteLine:destructive',
      'Document.FailHere:write',
    ])
  })

  it('preview is the dry run, execute is one transaction with a change set', async () => {
    const doc = toyDocument('doc_a')
    const dry = await doc.registry.run(
      'Document.AppendLine',
      { text: 'c' },
      { ...run, docId: 'doc_a', dryRun: true },
    )
    expect(dry.preview.summary).toEqual(['append "c"'])
    expect(doc.lines()).toEqual(['a', 'b'])
    const done = await doc.registry.run(
      'Document.AppendLine',
      { text: 'c' },
      { ...run, docId: 'doc_a' },
    )
    expect(doc.lines()).toEqual(['a', 'b', 'c'])
    expect(done.changeSet).toMatchObject({ action: 'Document.AppendLine', source: 'workflow' })
  })
})

describe('saga (Transaction acceptance)', () => {
  it('a failure halfway restores every touched document to its checkpoint', async () => {
    const a = toyDocument('doc_a')
    const b = toyDocument('doc_b')
    const participants = new Map<string, SagaParticipant>([
      ['doc_a', a.participant],
      ['doc_b', b.participant],
    ])
    const result = await runSaga(
      [
        { docId: 'doc_a', action: 'Document.AppendLine', input: { text: 'a1' } },
        { docId: 'doc_b', action: 'Document.DeleteLine', input: { index: 0 } },
        { docId: 'doc_a', action: 'Document.AppendLine', input: { text: 'a2' } },
        { docId: 'doc_b', action: 'Document.FailHere', input: {} },
        { docId: 'doc_a', action: 'Document.AppendLine', input: { text: 'never' } },
      ],
      participants,
      run,
    )
    expect(result.ok).toBe(false)
    expect(result.failedAt).toBe(3)
    expect(result.error).toBe('disk full')
    expect(a.lines()).toEqual(['a', 'b'])
    expect(b.lines()).toEqual(['a', 'b'])
    expect(result.rolledBack).toEqual(['doc_b', 'doc_a'])
    expect(result.journal.map((e) => e.kind)).toEqual([
      'checkpoint',
      'step',
      'checkpoint',
      'step',
      'step',
      'failed',
      'restored',
      'restored',
    ])
    // the steps that ran share the saga's txId
    expect(new Set(result.changeSets.map((c) => c.txId))).toEqual(new Set([result.txId]))
  })

  it('a clean saga keeps every change; preview changes nothing', async () => {
    const a = toyDocument('doc_a')
    const participants = new Map([['doc_a', a.participant]])
    const steps = [
      { docId: 'doc_a', action: 'Document.AppendLine', input: { text: 'x' } },
      { docId: 'doc_a', action: 'Document.DeleteLine', input: { index: 9 } },
    ]
    const preview = await previewSaga(steps, participants, run)
    expect(preview[0]!.preview!.summary).toEqual(['append "x"'])
    expect(preview[1]!.error).toBe('no line 9')
    expect(a.lines()).toEqual(['a', 'b'])
    const ok = await runSaga([steps[0]!], participants, run)
    expect(ok.ok).toBe(true)
    expect(a.lines()).toEqual(['a', 'b', 'x'])
  })

  it('a step in a document that is not open fails the saga', async () => {
    const a = toyDocument('doc_a')
    const result = await runSaga(
      [
        { docId: 'doc_a', action: 'Document.AppendLine', input: { text: 'x' } },
        { docId: 'doc_closed', action: 'Document.AppendLine', input: { text: 'y' } },
      ],
      new Map([['doc_a', a.participant]]),
      run,
    )
    expect(result.ok).toBe(false)
    expect(a.lines()).toEqual(['a', 'b'])
  })
})

describe('refusals and policy', () => {
  it('refuses actions outside the catalog, beyond permissions, or planned on a stale catalog', async () => {
    const doc = toyDocument('doc_a')
    const ai = { caller: 'ai' as const, docId: 'doc_a', permissions: defaultPermissions('ai') }
    await expect(doc.registry.run('File.DeleteEverything', {}, ai)).rejects.toMatchObject({
      code: 'unknown_action',
    })
    await expect(doc.registry.run('Document.DeleteLine', { index: 0 }, ai)).rejects.toMatchObject({
      code: 'permission_denied',
    })
    await expect(
      doc.registry.run('Document.AppendLine', { text: 'x', extra: true }, ai),
    ).rejects.toMatchObject({ code: 'invalid_input' })
    const fingerprint = doc.registry.fingerprint()
    expect(fingerprint).toMatch(/^[0-9a-f]{12}$/)
    await doc.registry.run('Document.AppendLine', { text: 'x' }, { ...ai, fingerprint })
    doc.registry.register({
      name: 'Document.NewThing',
      group: 'Document',
      summary: 'added later',
      input: z.object({}),
      effect: 'write',
      preview: () => ({ objects: 0, files: [], summary: [], warnings: [] }),
      execute: () => null,
    })
    const stale = doc.registry.run('Document.AppendLine', { text: 'y' }, { ...ai, fingerprint })
    await expect(stale).rejects.toBeInstanceOf(ActionError)
    await expect(stale).rejects.toMatchObject({ code: 'stale_catalog' })
  })

  it('confirms destructive/external always, and anything over the bulk threshold', () => {
    const report = (objects: number) => ({ objects, files: [], summary: [], warnings: [] })
    expect(needsConfirmation('write', report(10))).toBe(false)
    expect(needsConfirmation('write', report(51))).toBe(true)
    expect(needsConfirmation('bulk', report(10), { bulkThreshold: 5 })).toBe(true)
    expect(needsConfirmation('destructive', report(0))).toBe(true)
    expect(needsConfirmation('external', report(0))).toBe(true)
    expect(needsConfirmation('write', report(200), { bulkThreshold: 500 })).toBe(false)
  })
})

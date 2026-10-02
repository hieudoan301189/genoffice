/**
 * P4 in Slides: every pptx-ops op is a Presentation.* action (preview = the
 * executor's plan, execute = one transaction); a saga checkpoint is a deck
 * snapshot and restoring it is itself undoable.
 */
import { describe, expect, it, vi } from 'vitest'
import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { opNames, runTxn } from '@genoffice/pptx-ops'
import { defaultPermissions, runSaga, type SagaParticipant } from '@genoffice/dvh-actions'
import type { Session } from '../src/main/session-state'
import { pushHistory } from '../src/main/session-state'
import { createSlidesDvhActions, slidesSagaParticipant } from '../src/main/dvh-actions'

vi.mock('electron', () => ({ BrowserWindow: { getFocusedWindow: () => null } }))
vi.mock('../src/main/fonts', () => ({ createSystemFontMetrics: () => ({}) }))

async function setup() {
  const session: Session = {
    path: '',
    opened: await openPptx(await createBlankPptx()),
    fitWidthPx: 1280,
    undoStack: [],
    redoStack: [],
  } as Session
  const host = {
    session: () => session,
    // the essentials of applySessionTxn: dry runs plan, applies push one snapshot
    applyTxn: (s: Session, req: { ops: unknown[]; dryRun?: boolean }) => {
      const ops = req.ops as Parameters<typeof runTxn>[1]['ops']
      if (req.dryRun) {
        const r = runTxn(s.opened, { ops, dryRun: true })
        return { applied: false, dryRun: true, plan: r.plan ?? [], failures: r.failures }
      }
      pushHistory(s)
      const r = runTxn(s.opened, { ops })
      if (!r.applied) s.undoStack.pop()
      return {
        applied: r.applied,
        failures: r.failures,
        records: (r.records ?? []).map((rec) => ({ op: rec.op.op })),
      }
    },
  }
  const registry = createSlidesDvhActions(host)
  return { session, host, registry }
}

const ui = {
  caller: 'ui' as const,
  docId: 'doc_p',
  permissions: defaultPermissions('ui'),
  confirm: () => true,
}

describe('Presentation.* adapter', () => {
  it('lists every pptx-ops op; preview plans without changing the deck', async () => {
    const { session, registry } = await setup()
    expect(registry.catalog().length).toBe(opNames().length)
    expect(registry.has('Presentation.DuplicateSlide')).toBe(true)
    expect(registry.get('Presentation.DeleteSlide')!.effect).toBe('destructive')
    const slides = session.opened.deck.slides.length
    const dry = await registry.run(
      'Presentation.DuplicateSlide',
      { target: { slide: 0 } },
      { ...ui, dryRun: true },
    )
    expect(dry.preview.summary.length).toBeGreaterThan(0)
    expect(session.opened.deck.slides.length).toBe(slides)
    await registry.run('Presentation.DuplicateSlide', { target: { slide: 0 } }, ui)
    expect(session.opened.deck.slides.length).toBe(slides + 1)
    expect(session.undoStack.length).toBe(1)
    await expect(
      registry.run('Presentation.DuplicateSlide', { target: { slide: 99 } }, ui),
    ).rejects.toThrow()
  })

  it('a failing saga restores the deck, and the restore can be undone', async () => {
    const { session, host, registry } = await setup()
    const slides = session.opened.deck.slides.length
    const failing: SagaParticipant = {
      docId: 'doc_x',
      run: async () => {
        throw new Error('boom')
      },
      checkpoint: async () => null,
      restore: async () => {},
    }
    const result = await runSaga(
      [
        { docId: 'doc_p', action: 'Presentation.DuplicateSlide', input: { target: { slide: 0 } } },
        { docId: 'doc_x', action: 'Document.SetFont', input: {} },
      ],
      new Map([
        ['doc_p', slidesSagaParticipant(host, registry, 'doc_p')],
        ['doc_x', failing],
      ]),
      { caller: 'workflow', permissions: defaultPermissions('ui'), confirm: () => true },
    )
    expect(result.ok).toBe(false)
    expect(session.opened.deck.slides.length).toBe(slides)
    // the duplicate's snapshot plus the restore's own snapshot
    expect(session.undoStack.length).toBe(2)
  })
})

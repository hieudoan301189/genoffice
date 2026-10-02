// DVH Action Core for Slides (P4 adapter). Every pptx-ops op is a
// Presentation.* action: the registry validates the op, preview is the
// executor's dry run (runTxn plan), execute is one atomic transaction with one
// undo snapshot — the same path as the AI batch surface (applySessionTxn).
// A saga checkpoint is a history snapshot of the deck; restoring pushes the
// current deck onto the undo stack first, so the rollback itself can be undone.
import {
  ActionRegistry,
  localParticipant,
  registerOpFamily,
  type SagaParticipant,
} from '@genoffice/dvh-actions'
import { opNames, opUsage } from '@genoffice/pptx-ops'
import type { ApplyTxnOp, ApplyTxnResult } from '../shared/ipc'
import {
  pushHistory,
  restoreSnapshot,
  scheduleDeckBroadcast,
  takeSnapshot,
  type Session,
} from './session-state'

export interface SlidesActionHost {
  session(): Session | null
  /** applySessionTxn of slides-main (kept injectable for tests) */
  applyTxn(session: Session, request: ApplyTxnOp): ApplyTxnResult | null
}

export function createSlidesDvhActions(host: SlidesActionHost): ActionRegistry {
  const registry = new ActionRegistry()
  registerOpFamily(registry, {
    group: 'Presentation',
    entries: opNames().map((op) => ({ op, summary: opUsage(op)?.split('\n')[0] ?? op })),
    run: async (op, dryRun) => {
      const session = host.session()
      if (!session) throw new Error('no presentation is open')
      const result = host.applyTxn(session, { ops: [op], dryRun })
      if (!result) throw new Error('no presentation is open')
      if (result.failures?.length) {
        throw new Error(result.failures.map((f) => f.error).join('\n'))
      }
      if (!dryRun && !result.applied) throw new Error('the edit was not applied')
      const records = result.records ?? []
      return {
        summary: dryRun
          ? (result.plan ?? [])
          : records.map((r) => `${r.op}${r.target ? ` ${r.target}` : ''}`),
        objects: Math.max(1, records.length),
        output: { applied: result.applied, records },
        changes: records.map((r) => ({
          objectId: r.target ?? r.op,
          path: r.op,
          before: null,
          after: r.created ?? null,
        })),
      }
    },
  })
  return registry
}

export function slidesSagaParticipant(
  host: SlidesActionHost,
  registry: ActionRegistry,
  docId: string,
): SagaParticipant {
  return localParticipant({
    docId,
    registry,
    checkpoint: () => {
      const session = host.session()
      if (!session) throw new Error('no presentation is open')
      return takeSnapshot(session)
    },
    restore: (checkpoint) => {
      const session = host.session()
      if (!session) throw new Error('no presentation is open')
      pushHistory(session)
      restoreSnapshot(session, checkpoint as ReturnType<typeof takeSnapshot>)
      scheduleDeckBroadcast(session)
    },
  })
}

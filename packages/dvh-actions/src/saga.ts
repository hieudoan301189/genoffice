/**
 * Cross-document transactions (P4): inside one app an action is already one
 * undo item. Steps that span documents run as a saga — each document takes a
 * checkpoint before its first step; when a step fails, every document touched
 * so far is restored to its checkpoint, newest first, and the journal records
 * what ran, what failed and what was rolled back.
 *
 * Participants are async so the same saga runs in one renderer (tests, a
 * workflow inside a document) or in the shell across tabs (MCP `dvh_execute`).
 */

import type { ChangeSet } from '@genoffice/dvh-model'
import type { ActionCaller, ActionEffect, ConfirmPolicy, PreviewReport, RunResult } from './index'

export interface SagaParticipant {
  readonly docId: string
  /** a label for the journal (file name or tab title) */
  readonly label?: string
  /** runs one action in this document (ActionRegistry.run, or a bridge to it) */
  run(name: string, input: unknown, options: SagaRunOptions): Promise<RunResult>
  /** captures the document's state before the saga first changes it */
  checkpoint(): Promise<unknown>
  /** puts the document back to a checkpoint (one undo item per restore, or a reload) */
  restore(checkpoint: unknown): Promise<void>
  /** drops a checkpoint the saga no longer needs (it succeeded) */
  release?(checkpoint: unknown): Promise<void>
}

export interface SagaRunOptions {
  readonly caller: ActionCaller
  readonly txId: string
  readonly dryRun?: boolean
  readonly permissions: ReadonlySet<ActionEffect>
  readonly policy?: ConfirmPolicy
  readonly confirm?: (name: string, preview: PreviewReport) => Promise<boolean> | boolean
}

export interface SagaStep {
  /** the document the step runs in */
  readonly docId: string
  readonly action: string
  readonly input: unknown
}

export type SagaLogEntry =
  | { readonly at: string; readonly kind: 'checkpoint'; readonly docId: string }
  | {
      readonly at: string
      readonly kind: 'step'
      readonly index: number
      readonly docId: string
      readonly action: string
      readonly changes: number
    }
  | {
      readonly at: string
      readonly kind: 'failed'
      readonly index: number
      readonly docId: string
      readonly action: string
      readonly error: string
    }
  | { readonly at: string; readonly kind: 'restored'; readonly docId: string }
  | {
      readonly at: string
      readonly kind: 'restore-failed'
      readonly docId: string
      readonly error: string
    }

export interface SagaResult {
  readonly ok: boolean
  readonly txId: string
  /** one per step that ran to completion */
  readonly results: readonly RunResult[]
  readonly changeSets: readonly ChangeSet[]
  readonly failedAt?: number
  readonly error?: string
  /** documents put back to their checkpoint after a failure */
  readonly rolledBack: readonly string[]
  readonly journal: readonly SagaLogEntry[]
}

export interface SagaOptions extends Omit<SagaRunOptions, 'txId' | 'dryRun'> {
  readonly txId?: string
  readonly now?: () => string
}

let sagaSeq = 0

/** Runs the steps in order as one transaction across documents. */
export async function runSaga(
  steps: readonly SagaStep[],
  participants: ReadonlyMap<string, SagaParticipant>,
  options: SagaOptions,
): Promise<SagaResult> {
  const now = options.now ?? (() => new Date().toISOString())
  const txId = options.txId ?? `tx_${Date.now().toString(36)}${(sagaSeq++).toString(36)}`
  const journal: SagaLogEntry[] = []
  const checkpoints = new Map<string, unknown>()
  /** documents in the order they were first touched */
  const touched: string[] = []
  const results: RunResult[] = []
  const changeSets: ChangeSet[] = []

  for (const [index, step] of steps.entries()) {
    const participant = participants.get(step.docId)
    try {
      if (!participant) throw new Error(`no open document ${step.docId}`)
      if (!checkpoints.has(step.docId)) {
        checkpoints.set(step.docId, await participant.checkpoint())
        touched.push(step.docId)
        journal.push({ at: now(), kind: 'checkpoint', docId: step.docId })
      }
      const result = await participant.run(step.action, step.input, { ...options, txId })
      results.push(result)
      if (result.changeSet) changeSets.push(result.changeSet)
      journal.push({
        at: now(),
        kind: 'step',
        index,
        docId: step.docId,
        action: step.action,
        changes: result.changeSet?.changes.length ?? 0,
      })
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      journal.push({
        at: now(),
        kind: 'failed',
        index,
        docId: step.docId,
        action: step.action,
        error: message,
      })
      const rolledBack: string[] = []
      for (const docId of [...touched].reverse()) {
        try {
          await participants.get(docId)!.restore(checkpoints.get(docId))
          rolledBack.push(docId)
          journal.push({ at: now(), kind: 'restored', docId })
        } catch (restoreError) {
          journal.push({
            at: now(),
            kind: 'restore-failed',
            docId,
            error: restoreError instanceof Error ? restoreError.message : String(restoreError),
          })
        }
      }
      return {
        ok: false,
        txId,
        results,
        changeSets,
        failedAt: index,
        error: message,
        rolledBack,
        journal,
      }
    }
  }
  for (const docId of touched) {
    try {
      await participants.get(docId)!.release?.(checkpoints.get(docId))
    } catch {
      // a checkpoint that cannot be released expires on its own
    }
  }
  return { ok: true, txId, results, changeSets, rolledBack: [], journal }
}

/** Previews every step without changing anything (each document's dry run). */
export async function previewSaga(
  steps: readonly SagaStep[],
  participants: ReadonlyMap<string, SagaParticipant>,
  options: SagaOptions,
): Promise<{ step: SagaStep; preview?: PreviewReport; error?: string }[]> {
  const txId = options.txId ?? 'preview'
  const out: { step: SagaStep; preview?: PreviewReport; error?: string }[] = []
  for (const step of steps) {
    const participant = participants.get(step.docId)
    if (!participant) {
      out.push({ step, error: `no open document ${step.docId}` })
      continue
    }
    try {
      const result = await participant.run(step.action, step.input, {
        ...options,
        txId,
        dryRun: true,
      })
      out.push({ step, preview: result.preview })
    } catch (error) {
      out.push({ step, error: error instanceof Error ? error.message : String(error) })
    }
  }
  return out
}

/** A participant over a local registry, with app-supplied checkpoint/restore. */
export function localParticipant(options: {
  docId: string
  label?: string
  registry: { run(name: string, input: unknown, options: never): Promise<RunResult> }
  checkpoint(): Promise<unknown> | unknown
  restore(checkpoint: unknown): Promise<void> | void
}): SagaParticipant {
  return {
    docId: options.docId,
    ...(options.label ? { label: options.label } : {}),
    run: (name, input, run) =>
      options.registry.run(name, input, { ...run, docId: options.docId } as never),
    checkpoint: async () => options.checkpoint(),
    restore: async (checkpoint) => options.restore(checkpoint),
  }
}

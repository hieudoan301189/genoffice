/**
 * Workflows in Docs (P8): one action registry for the UI, a recorder on its
 * event stream, and the runner the Workflow section uses.
 *
 * UI commands reach the recorder two ways: commands that are catalog actions
 * run through the registry (`runUiAction`); commands that keep their own code
 * path announce the catalog action they amount to (`announceUi`). Commands
 * with no catalog equivalent yet tell the recorder they cannot be recorded.
 */
import { defaultPermissions, type ActionRegistry, type PreviewReport } from '@genoffice/dvh-actions'
import type { Scalar } from '@genoffice/dvh-model'
import {
  WorkflowRecorder,
  runWorkflow,
  type Json,
  type PauseDecision,
  type PauseInfo,
  type Workflow,
  type WorkflowRunResult,
} from '@genoffice/dvh-workflow'
import { createDocsDvhActions, docsSagaParticipant, type DocsActionHost } from './dvh-actions'
import { activeDvhDocs } from './dvh-smart-data'

let shared: { registry: ActionRegistry; host: DocsActionHost } | null = null
let selection: Record<string, Scalar> | null = null
const listeners = new Set<() => void>()

/** The recorder of the open document's UI actions. */
export const docsRecorder = new WorkflowRecorder({
  context: () => (selection ? { record: { values: selection } } : undefined),
})

/** Creates the UI registry once and attaches the recorder to it. */
export function installDocsUiActions(host: DocsActionHost): ActionRegistry {
  if (shared) return shared.registry
  const registry = createDocsDvhActions(host)
  docsRecorder.attach(registry)
  registry.subscribe(() => notify())
  shared = { registry, host }
  return registry
}

export function docsUiRegistry(): ActionRegistry | null {
  return shared?.registry ?? null
}

/** Test hook: forget the registry (a fresh one is created on the next install). */
export function resetDocsUiActions(): void {
  shared = null
  selection = null
  docsRecorder.stop()
  docsRecorder.clear()
}

/** The record the user has selected (its values become workflow parameters when recorded). */
export function setRecordSelection(values: Record<string, Scalar> | null): void {
  selection = values
}

export function onRecorderChange(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

function notify(): void {
  for (const listener of [...listeners]) listener()
}

const docIdNow = () => activeDvhDocs()?.model?.docId ?? 'doc'

/** Runs a catalog action as the user (recorded while recording). */
export async function runUiAction(name: string, input: unknown): Promise<unknown> {
  if (!shared) throw new Error('the action registry is not ready')
  const result = await shared.registry.run(name, input, {
    caller: 'ui',
    docId: docIdNow(),
    permissions: defaultPermissions('ui'),
    confirm: (action, preview) => shared?.host.confirm?.(action, preview) ?? false,
  })
  return result.output
}

/** A UI command that ran its own way but means exactly this catalog action. */
export function announceUi(name: string, input: unknown): void {
  shared?.registry.announceUi(name, input, docIdNow())
}

/** A UI command with no catalog action yet: the recording warns instead of silently missing it. */
export function noteUnrecordable(command: string): void {
  if (!docsRecorder.recording) return
  docsRecorder.unrecordable(command)
  notify()
}

export function startRecording(): void {
  docsRecorder.clear()
  docsRecorder.start()
  notify()
}

export function stopRecording(): void {
  docsRecorder.stop()
  notify()
}

/** Runs a workflow in the open document as the user, optionally step by step. */
export async function runDocsWorkflow(
  workflow: Workflow,
  options: {
    params?: Record<string, Json>
    dryRun?: boolean
    stepMode?: boolean
    breakpoints?: ReadonlySet<string>
    onPause?: (pause: PauseInfo) => Promise<PauseDecision>
    confirm?: (name: string, preview: PreviewReport) => Promise<boolean> | boolean
  } = {},
): Promise<WorkflowRunResult> {
  if (!shared) throw new Error('the action registry is not ready')
  const docId = docIdNow()
  return runWorkflow(workflow, {
    participants: new Map([[docId, docsSagaParticipant(shared.host, shared.registry, docId)]]),
    docId,
    // the steps run as `workflow` (never recorded) with what the user may do
    caller: 'workflow',
    permissions: defaultPermissions('ui'),
    ...(options.params ? { params: options.params } : {}),
    ...(options.dryRun ? { dryRun: true } : {}),
    ...(options.stepMode ? { stepMode: true } : {}),
    ...(options.breakpoints ? { breakpoints: options.breakpoints } : {}),
    ...(options.onPause ? { onPause: options.onPause } : {}),
    confirm: options.confirm ?? ((name, preview) => shared?.host.confirm?.(name, preview) ?? false),
  })
}

/**
 * Running a script (P9): compile, check against the catalog, then interpret
 * the workflow model with the P8 engine as caller `script`. There is no
 * `eval` and nothing reaches Node, the file system or the network: the only
 * way out is a catalog action, which checks its permission (scripts get read +
 * write unless the user grants more) — and steps and time are limited.
 */

import { defaultPermissions, type ActionEffect } from '@genoffice/dvh-actions'
import {
  runWorkflow,
  type RunWorkflowOptions,
  type WorkflowRunResult,
} from '@genoffice/dvh-workflow'
import { checkScript, hasErrors, type Diagnostic, type ScriptCatalogEntry } from './check'

export const SCRIPT_LIMITS = { maxSteps: 10_000, timeoutMs: 30_000 } as const

export type ScriptRunOutcome =
  | { readonly ok: false; readonly diagnostics: readonly Diagnostic[] }
  | {
      readonly ok: true
      readonly result: WorkflowRunResult
      /** source line of each step path (the debugger shows the paused line) */
      readonly lines: Readonly<Record<string, number>>
      readonly diagnostics: readonly Diagnostic[]
    }

export async function runScript(
  source: string,
  catalog: readonly ScriptCatalogEntry[],
  options: Omit<RunWorkflowOptions, 'caller' | 'permissions'> & {
    readonly permissions?: ReadonlySet<ActionEffect>
  },
): Promise<ScriptRunOutcome> {
  const checked = checkScript(source, catalog)
  if (!checked.workflow || hasErrors(checked.diagnostics))
    return { ok: false, diagnostics: checked.diagnostics }
  const result = await runWorkflow(checked.workflow, {
    maxSteps: SCRIPT_LIMITS.maxSteps,
    timeoutMs: SCRIPT_LIMITS.timeoutMs,
    ...options,
    caller: 'script',
    permissions: options.permissions ?? defaultPermissions('script'),
  })
  return { ok: true, result, lines: checked.lines, diagnostics: checked.diagnostics }
}

/** The source line a step path came from, for errors and the debugger. */
export function lineOfPath(
  lines: Readonly<Record<string, number>>,
  path: string | undefined,
): number | null {
  if (!path) return null
  let p = path
  for (;;) {
    if (lines[p] !== undefined) return lines[p]!
    const cut = p.lastIndexOf('.')
    if (cut < 0) return null
    p = p.slice(0, cut)
  }
}

/**
 * The Workflow Engine (P8): runs a workflow against open documents through
 * their saga participants (P4), so every action is validated, permission
 * checked, confirmed by policy and recorded with the run's txId.
 *
 * - A run is a transaction: an uncaught error puts every document back to the
 *   checkpoint taken before the run first changed it.
 * - `Transaction` steps nest: a failure inside one rolls back only its own
 *   changes, and a `Try` around it can carry on.
 * - Step-by-step runs and breakpoints pause before a step and ask the host
 *   whether to continue, step or stop; every step lands in the log.
 * - Limits on steps and time hold even inside `Try` (the sandbox of P9).
 */

import {
  defaultPermissions,
  type ActionCaller,
  type ActionEffect,
  type ConfirmPolicy,
  type PreviewReport,
  type SagaParticipant,
} from '@genoffice/dvh-actions'
import type { ChangeSet, Scalar } from '@genoffice/dvh-model'
import { evaluateTree, parseExpr, truthy, type ExprNode } from '@genoffice/dvh-template'
import { type Arg, type Json, type Step, type Workflow, type WorkflowParam } from './model'

export type LogEntry = {
  readonly at: string
  readonly path: string
  readonly kind:
    | 'action'
    | 'preview'
    | 'set'
    | 'branch'
    | 'loop'
    | 'log'
    | 'error'
    | 'caught'
    | 'rollback'
    | 'pause'
  readonly message: string
  readonly action?: string
  readonly docId?: string
  readonly objects?: number
}

export type PauseDecision = 'continue' | 'step' | 'stop'

export interface PauseInfo {
  readonly path: string
  readonly step: Step
  /** a snapshot of the variables in scope */
  readonly variables: Record<string, Json>
}

export interface RunWorkflowOptions {
  /** values of the workflow's parameters (defaults fill the rest) */
  readonly params?: Record<string, Json>
  /** the documents actions may run in */
  readonly participants: ReadonlyMap<string, SagaParticipant>
  /** where an action without `doc` runs */
  readonly docId: string
  readonly caller?: Extract<ActionCaller, 'workflow' | 'script' | 'ui' | 'ai'>
  readonly permissions?: ReadonlySet<ActionEffect>
  readonly policy?: ConfirmPolicy
  readonly confirm?: (name: string, preview: PreviewReport) => Promise<boolean> | boolean
  readonly txId?: string
  /** previews every action instead of running it; nothing changes */
  readonly dryRun?: boolean
  /** pause before the first step and after every step */
  readonly stepMode?: boolean
  /** step paths to pause at */
  readonly breakpoints?: ReadonlySet<string>
  readonly onPause?: (pause: PauseInfo) => Promise<PauseDecision> | PauseDecision
  readonly onLog?: (entry: LogEntry) => void
  /** steps executed before the run is stopped (default 10 000) */
  readonly maxSteps?: number
  /** wall-clock budget in milliseconds (default 120 000) */
  readonly timeoutMs?: number
  readonly signal?: AbortSignal
  readonly now?: () => string
}

export interface WorkflowRunResult {
  readonly ok: boolean
  readonly txId: string
  readonly stopped: boolean
  readonly error?: string
  /** the step that failed */
  readonly failedAt?: string
  readonly log: readonly LogEntry[]
  readonly changeSets: readonly ChangeSet[]
  readonly previews: readonly { path: string; action: string; preview: PreviewReport }[]
  /** documents put back to their checkpoint by the run's rollback */
  readonly rolledBack: readonly string[]
  readonly steps: number
  /** top-level variables at the end */
  readonly variables: Record<string, Json>
}

/** Errors no `Try` may catch: limits, a stop from the debugger, an abort. */
export class WorkflowHalt extends Error {
  constructor(
    message: string,
    readonly stopped = false,
  ) {
    super(message)
    this.name = 'WorkflowHalt'
  }
}

class StepError extends Error {
  constructor(
    message: string,
    readonly path: string,
  ) {
    super(message)
    this.name = 'StepError'
  }
}

let runSeq = 0

type Scope = Map<string, Json>[]

function resolvePath(scope: Scope, name: string): Json | undefined {
  const [head, ...rest] = name.split('.')
  let value: Json | undefined
  for (let i = scope.length - 1; i >= 0; i--) {
    if (scope[i]!.has(head!)) {
      value = scope[i]!.get(head!)
      break
    }
  }
  for (const key of rest) {
    if (value === undefined || value === null) return undefined
    if (Array.isArray(value)) {
      // VBA-like `.Count` of a list
      if (key === 'Count' || key === 'count') return value.length
      return undefined
    }
    if (typeof value !== 'object') return undefined
    value = Object.prototype.hasOwnProperty.call(value, key) ? value[key] : undefined
  }
  return value
}

const isScalar = (v: Json): v is Scalar =>
  v === null || typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean'

const treeCache = new Map<string, ExprNode>()
function treeOf(source: string): ExprNode {
  let tree = treeCache.get(source)
  if (!tree) {
    tree = parseExpr(source)
    if (treeCache.size > 2000) treeCache.clear()
    treeCache.set(source, tree)
  }
  return tree
}

/** Evaluates an expression in a scope; a bare name returns its value as is (a record, a list). */
export function evaluateIn(scope: Scope, source: string): Json {
  const tree = treeOf(source)
  if (tree.k === 'name') {
    const value = resolvePath(scope, tree.name)
    if (value === undefined) throw new Error(`unknown name ${tree.name}`)
    return structuredClone(value)
  }
  return evaluateTree(tree, {
    lookup: (name) => {
      const value = resolvePath(scope, name)
      if (value === undefined) return undefined
      if (!isScalar(value))
        throw new Error(`${name} is a ${Array.isArray(value) ? 'list' : 'record'}, not a value`)
      return value
    },
    column: (title) => {
      // `[Title]` reads the innermost loop record by key
      for (let i = scope.length - 1; i >= 0; i--) {
        const record = scope[i]!.get('$record')
        if (record && typeof record === 'object' && !Array.isArray(record)) {
          const value = record[title]
          if (value === undefined) return undefined
          return isScalar(value) ? value : undefined
        }
      }
      return undefined
    },
  })
}

export function evaluateArg(scope: Scope, arg: Arg): Json {
  return 'value' in arg ? structuredClone(arg.value) : evaluateIn(scope, arg.expr)
}

/** Fills parameter values from defaults and coerces them to their declared type. */
export function bindParams(
  params: readonly WorkflowParam[],
  given: Record<string, Json> = {},
): Map<string, Json> {
  const out = new Map<string, Json>()
  for (const p of params) {
    let value = Object.prototype.hasOwnProperty.call(given, p.name) ? given[p.name] : p.default
    if (value === undefined) throw new Error(`parameter ${p.name} has no value`)
    switch (p.type) {
      case 'number':
        if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value)))
          value = Number(value)
        if (typeof value !== 'number') throw new Error(`parameter ${p.name} must be a number`)
        break
      case 'boolean':
        if (value === 'true' || value === 'TRUE') value = true
        if (value === 'false' || value === 'FALSE') value = false
        if (typeof value !== 'boolean') throw new Error(`parameter ${p.name} must be true or false`)
        break
      case 'text':
        if (typeof value === 'number' || typeof value === 'boolean') value = String(value)
        if (typeof value !== 'string' && value !== null)
          throw new Error(`parameter ${p.name} must be text`)
        break
      case 'record':
        if (!value || typeof value !== 'object' || Array.isArray(value))
          throw new Error(`parameter ${p.name} must be a record`)
        break
      case 'list':
        if (!Array.isArray(value)) throw new Error(`parameter ${p.name} must be a list`)
        break
      case 'any':
        break
    }
    out.set(p.name, value)
  }
  return out
}

interface TxFrame {
  readonly checkpoints: Map<string, unknown>
  readonly order: string[]
}

/** Runs a workflow; never throws — the result says how it went. */
export async function runWorkflow(
  workflow: Workflow,
  options: RunWorkflowOptions,
): Promise<WorkflowRunResult> {
  const now = options.now ?? (() => new Date().toISOString())
  const caller = options.caller ?? 'workflow'
  const permissions = options.permissions ?? defaultPermissions(caller)
  const txId = options.txId ?? `tx_${Date.now().toString(36)}w${(runSeq++).toString(36)}`
  const maxSteps = options.maxSteps ?? 10_000
  const deadline = Date.now() + (options.timeoutMs ?? 120_000)
  const log: LogEntry[] = []
  const changeSets: ChangeSet[] = []
  const previews: { path: string; action: string; preview: PreviewReport }[] = []
  const frames: TxFrame[] = []
  const rolledBack: string[] = []
  let steps = 0
  let stepMode = options.stepMode ?? false

  const write = (entry: Omit<LogEntry, 'at'>) => {
    const full = { at: now(), ...entry }
    log.push(full)
    options.onLog?.(full)
  }

  const snapshot = (scope: Scope): Record<string, Json> => {
    const out: Record<string, Json> = {}
    for (const frame of scope) for (const [k, v] of frame) if (!k.startsWith('$')) out[k] = v
    return structuredClone(out)
  }

  const assign = (scope: Scope, name: string, value: Json) => {
    for (let i = scope.length - 1; i >= 0; i--) {
      if (scope[i]!.has(name)) {
        scope[i]!.set(name, value)
        return
      }
    }
    // as in VBA, a new variable belongs to the whole run, not to the block it was first set in
    scope[0]!.set(name, value)
  }

  const rollback = async (frame: TxFrame, path: string) => {
    for (const docId of [...frame.order].reverse()) {
      try {
        await options.participants.get(docId)!.restore(frame.checkpoints.get(docId))
        if (!rolledBack.includes(docId)) rolledBack.push(docId)
        write({ path, kind: 'rollback', message: `restored ${docId}`, docId })
      } catch (error) {
        write({
          path,
          kind: 'error',
          message: `could not restore ${docId}: ${error instanceof Error ? error.message : String(error)}`,
          docId,
        })
      }
    }
  }

  const inTransaction = async (path: string, run: () => Promise<void>) => {
    const frame: TxFrame = { checkpoints: new Map(), order: [] }
    frames.push(frame)
    try {
      await run()
      if (frames.length === 1) {
        // the run succeeded: its checkpoints are no longer needed
        for (const docId of frame.order) {
          try {
            await options.participants.get(docId)!.release?.(frame.checkpoints.get(docId))
          } catch {
            // a checkpoint that cannot be released expires on its own
          }
        }
      }
    } catch (error) {
      if (!options.dryRun) await rollback(frame, path)
      throw error
    } finally {
      frames.pop()
    }
  }

  const before = async (path: string, step: Step, scope: Scope) => {
    if (options.signal?.aborted) throw new WorkflowHalt('the run was cancelled', true)
    if (++steps > maxSteps) throw new WorkflowHalt(`the run exceeded ${maxSteps} steps`)
    if (Date.now() > deadline) throw new WorkflowHalt('the run took too long')
    if (stepMode || options.breakpoints?.has(path)) {
      write({ path, kind: 'pause', message: step.kind })
      const decision =
        (await options.onPause?.({ path, step, variables: snapshot(scope) })) ?? 'continue'
      if (decision === 'stop') throw new WorkflowHalt('stopped in the debugger', true)
      stepMode = decision === 'step'
    }
  }

  const runAction = async (step: Extract<Step, { kind: 'action' }>, path: string, scope: Scope) => {
    const docRef = step.doc ? evaluateArg(scope, step.doc) : options.docId
    if (typeof docRef !== 'string') throw new Error('the document of a step must be a text id')
    const participant = options.participants.get(docRef)
    if (!participant) throw new Error(`no open document ${docRef}`)
    const input: Record<string, Json> = {}
    for (const [key, arg] of Object.entries(step.args)) input[key] = evaluateArg(scope, arg)
    if (!options.dryRun) {
      for (const frame of frames) {
        if (frame.checkpoints.has(docRef)) continue
        frame.checkpoints.set(docRef, await participant.checkpoint())
        frame.order.push(docRef)
      }
    }
    const result = await participant.run(step.action, input, {
      caller,
      txId,
      permissions,
      ...(options.policy ? { policy: options.policy } : {}),
      ...(options.confirm ? { confirm: options.confirm } : {}),
      ...(options.dryRun ? { dryRun: true } : {}),
    })
    if (options.dryRun) {
      previews.push({ path, action: step.action, preview: result.preview })
      write({
        path,
        kind: 'preview',
        message: result.preview.summary.join('; '),
        action: step.action,
        docId: docRef,
        objects: result.preview.objects,
      })
    } else {
      if (result.changeSet) changeSets.push(result.changeSet)
      write({
        path,
        kind: 'action',
        message: result.preview.summary.join('; '),
        action: step.action,
        docId: docRef,
        objects: result.changeSet?.changes.length ?? 0,
      })
    }
    if (step.assign) {
      const output = result.output === undefined || options.dryRun ? null : result.output
      assign(scope, step.assign, structuredClone(output as Json))
    }
  }

  const runList = async (list: readonly Step[], base: string, scope: Scope): Promise<void> => {
    for (const [i, step] of list.entries()) {
      const path = base ? `${base}.${i}` : String(i)
      await before(path, step, scope)
      try {
        await runStep(step, path, scope)
      } catch (error) {
        if (error instanceof WorkflowHalt || error instanceof StepError) throw error
        throw new StepError(error instanceof Error ? error.message : String(error), path)
      }
    }
  }

  const runStep = async (step: Step, path: string, scope: Scope): Promise<void> => {
    switch (step.kind) {
      case 'action':
        return runAction(step, path, scope)
      case 'set': {
        const value = evaluateArg(scope, step.value)
        assign(scope, step.variable, value)
        write({
          path,
          kind: 'set',
          message: `${step.variable} = ${JSON.stringify(value)?.slice(0, 80)}`,
        })
        return
      }
      case 'log': {
        const value = evaluateArg(scope, step.message)
        write({
          path,
          kind: 'log',
          message: typeof value === 'string' ? value : JSON.stringify(value),
        })
        return
      }
      case 'if': {
        for (const [i, branch] of step.branches.entries()) {
          if (truthy(toScalar(evaluateIn(scope, branch.when)))) {
            write({ path, kind: 'branch', message: `branch ${i + 1}` })
            return runList(branch.steps, `${path}.then${i}`, [...scope, new Map()])
          }
        }
        if (step.else) {
          write({ path, kind: 'branch', message: 'else' })
          return runList(step.else, `${path}.else`, [...scope, new Map()])
        }
        return
      }
      case 'forEach': {
        const items = evaluateArg(scope, step.items)
        if (!Array.isArray(items)) throw new Error(`ForEach ${step.variable}: not a list`)
        let index = 0
        for (const item of items) {
          index++
          const frame = new Map<string, Json>([
            [step.variable, item],
            ['INDEX', index],
            ...(item && typeof item === 'object' && !Array.isArray(item)
              ? [['$record', item] as [string, Json]]
              : []),
          ])
          const inner = [...scope, frame]
          if (step.where && !truthy(toScalar(evaluateIn(inner, step.where)))) continue
          write({ path, kind: 'loop', message: `${step.variable} #${index}` })
          await runList(step.steps, `${path}.body`, inner)
        }
        return
      }
      case 'try': {
        try {
          await runList(step.steps, `${path}.body`, [...scope, new Map()])
        } catch (error) {
          if (error instanceof WorkflowHalt) throw error
          const message = error instanceof Error ? error.message : String(error)
          write({ path, kind: 'caught', message })
          const frame = new Map<string, Json>([
            [
              step.errorVariable ?? 'error',
              { message, step: error instanceof StepError ? error.path : path },
            ],
          ])
          await runList(step.catch, `${path}.catch`, [...scope, frame])
        }
        return
      }
      case 'transaction':
        return inTransaction(path, () => runList(step.steps, `${path}.body`, [...scope, new Map()]))
    }
  }

  let top: Map<string, Json>
  try {
    top = bindParams(workflow.params, options.params)
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    write({ path: '', kind: 'error', message })
    return {
      ok: false,
      txId,
      stopped: false,
      error: message,
      log,
      changeSets,
      previews,
      rolledBack,
      steps,
      variables: {},
    }
  }
  const scope: Scope = [top]
  try {
    await inTransaction('', () => runList(workflow.steps, '', scope))
    return {
      ok: true,
      txId,
      stopped: false,
      log,
      changeSets,
      previews,
      rolledBack,
      steps,
      variables: snapshot(scope),
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    const failedAt = error instanceof StepError ? error.path : undefined
    write({ path: failedAt ?? '', kind: 'error', message })
    return {
      ok: false,
      txId,
      stopped: error instanceof WorkflowHalt && error.stopped,
      error: message,
      ...(failedAt !== undefined ? { failedAt } : {}),
      log,
      changeSets,
      previews,
      rolledBack,
      steps,
      variables: snapshot(scope),
    }
  }
}

function toScalar(value: Json): Scalar {
  if (isScalar(value)) return value
  return Array.isArray(value) ? value.length > 0 : Object.keys(value).length > 0
}

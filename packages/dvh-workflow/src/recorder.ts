/**
 * Record Anything (P8): the recorder listens to the Action Core event stream
 * (actions the user ran from the UI), never to mouse or keyboard, so a
 * recorded workflow keeps working when the ribbon or panels are rearranged.
 *
 * Normalisation:
 * - selection-only actions (group `UI`, or effect `read`) are dropped;
 * - consecutive `Set…` actions on the same target keep only the last value;
 * - an undo cancels the step it undid, a redo brings it back.
 *
 * Parameter inference: a top-level argument equal to a value of the record
 * selected when the action ran becomes `record.<key>`, and the workflow gets a
 * `record` parameter (defaulting to the recorded record), so it can run again
 * for another record — or for each record of a list.
 */

import type { ActionEvent } from '@genoffice/dvh-actions'
import type { Scalar } from '@genoffice/dvh-model'
import {
  IDENT,
  newWorkflowId,
  type Arg,
  type Json,
  type Step,
  type Workflow,
  type WorkflowParam,
} from './model'

/** What the UI had selected when an action ran. */
export interface RecordContext {
  readonly record?: { readonly values: Readonly<Record<string, Scalar>> }
}

export interface RecordedStep {
  readonly action: string
  readonly docId: string
  readonly input: Record<string, Json>
  /** the selected record when it ran */
  readonly record?: Readonly<Record<string, Scalar>>
}

export interface RecorderOptions {
  /** called for each event to learn the current selection */
  readonly context?: () => RecordContext | undefined
  /** extra noise filter on top of the built-in rules */
  readonly ignore?: (event: ActionEvent) => boolean
}

const UNDO = /\.Undo$/
const REDO = /\.Redo$/
const SETTER = /\.Set[A-Z]\w*$/
const VALUE_KEYS = new Set(['value', 'values', 'text'])

/** The part of a setter's input that names its target (everything but the value). */
function targetKey(step: RecordedStep): string | null {
  if (!SETTER.test(step.action)) return null
  const target = Object.entries(step.input)
    .filter(([k]) => !VALUE_KEYS.has(k))
    .sort(([a], [b]) => a.localeCompare(b))
  return `${step.docId}|${step.action}|${JSON.stringify(target)}`
}

export class WorkflowRecorder {
  private readonly recorded: RecordedStep[] = []
  private readonly undone: RecordedStep[] = []
  private readonly missed: string[] = []
  private active = false

  constructor(private readonly options: RecorderOptions = {}) {}

  get recording(): boolean {
    return this.active
  }

  start(): void {
    this.active = true
  }

  stop(): void {
    this.active = false
  }

  clear(): void {
    this.recorded.length = 0
    this.undone.length = 0
    this.missed.length = 0
  }

  /** Subscribes to a registry; only actions the user ran (`ui`) are recorded. */
  attach(registry: { subscribe(listener: (event: ActionEvent) => void): () => void }): () => void {
    return registry.subscribe((event) => this.record(event))
  }

  /** Takes one event from the stream (ignored while not recording). */
  record(event: ActionEvent, context = this.options.context?.()): void {
    if (!this.active || event.caller !== 'ui') return
    if (UNDO.test(event.name)) {
      const last = this.recorded.pop()
      if (last) this.undone.push(last)
      return
    }
    if (REDO.test(event.name)) {
      const again = this.undone.pop()
      if (again) this.recorded.push(again)
      return
    }
    if (event.group === 'UI' || event.effect === 'read') return
    if (this.options.ignore?.(event)) return
    const step: RecordedStep = {
      action: event.name,
      docId: event.docId,
      input: structuredClone((event.input ?? {}) as Record<string, Json>),
      ...(context?.record ? { record: { ...context.record.values } } : {}),
    }
    // a new action makes the undone ones unreachable, as in any undo stack
    this.undone.length = 0
    const key = targetKey(step)
    const previous = this.recorded.at(-1)
    if (key !== null && previous && targetKey(previous) === key) this.recorded.pop()
    this.recorded.push(step)
  }

  /**
   * A UI command that does not go through the Action Core yet: it cannot be
   * recorded, and the user is told so instead of getting a silently partial
   * workflow.
   */
  unrecordable(command: string): void {
    if (!this.active) return
    if (!this.missed.includes(command)) this.missed.push(command)
  }

  /** The commands used while recording that could not be recorded (each once). */
  unrecorded(): readonly string[] {
    return this.missed
  }

  steps(): readonly RecordedStep[] {
    return this.recorded
  }

  warnings(): readonly string[] {
    return this.missed.map((command) => `"${command}" cannot be recorded`)
  }

  /**
   * The recording as a workflow. With `forEach`, the steps run once per item
   * of a `records` list parameter (each item bound to `record`).
   */
  toWorkflow(options: {
    name: string
    id?: string
    catalog?: string
    forEach?: boolean
  }): Workflow {
    const docs = [...new Set(this.recorded.map((s) => s.docId))]
    const home = docs[0]
    let inferred: Record<string, Scalar> | null = null
    const steps: Step[] = this.recorded.map((recorded) => {
      const args: Record<string, Arg> = {}
      for (const [key, value] of Object.entries(recorded.input)) {
        const from = recorded.record ? recordKeyOf(recorded.record, value) : null
        if (from) {
          args[key] = { expr: `record.${from}` }
          inferred ??= { ...recorded.record! }
        } else {
          args[key] = { value }
        }
      }
      return {
        kind: 'action',
        action: recorded.action,
        args,
        ...(recorded.docId !== home ? { doc: { value: recorded.docId } } : {}),
      }
    })
    const params: WorkflowParam[] = []
    let body = steps
    if (inferred && options.forEach) {
      params.push({ name: 'records', type: 'list', default: [inferred] })
      body = [{ kind: 'forEach', variable: 'record', items: { expr: 'records' }, steps }]
    } else if (inferred) {
      params.push({ name: 'record', type: 'record', default: inferred })
    }
    return {
      format: 'dvh-workflow',
      version: 1,
      id: options.id ?? newWorkflowId(),
      name: options.name,
      params,
      steps: body,
      ...(options.catalog ? { catalog: options.catalog } : {}),
    }
  }
}

/** The key of the record value an argument equals (non-empty values only). */
function recordKeyOf(record: Readonly<Record<string, Scalar>>, value: Json): string | null {
  if (value === null || typeof value === 'object' || value === '' || typeof value === 'boolean')
    return null
  for (const [key, candidate] of Object.entries(record)) {
    if (!IDENT.test(key) || candidate === null || candidate === '') continue
    if (candidate === value || String(candidate) === String(value)) return key
  }
  return null
}

/** The rows of a collection as records keyed by column key (`record.qty`), for ForEach and parameters. */
export function recordsOf(collection: {
  readonly columns: readonly { readonly key: string }[]
  readonly rows: readonly (readonly Scalar[])[]
}): Record<string, Scalar>[] {
  return collection.rows.map((row) =>
    Object.fromEntries(collection.columns.map((c, i) => [c.key, row[i] ?? null])),
  )
}

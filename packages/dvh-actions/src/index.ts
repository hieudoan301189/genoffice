/**
 * DVH Action Core (ADR D6, D7): one action catalog shared by the UI, AI,
 * the recorder and DVH-Script. An action validates its input, declares its
 * effect class (the permission it needs), previews without side effects, and
 * reports what it changed as change sets of one transaction.
 */

import { z, type ZodType } from 'zod'
import type { ChangeSet } from '@genoffice/dvh-model'

export type ActionEffect = 'read' | 'write' | 'bulk' | 'destructive' | 'external'
export type ActionCaller = 'ui' | 'ai' | 'workflow' | 'script' | 'extension'
export type ActionGroup =
  'Data' | 'Table' | 'Document' | 'Spreadsheet' | 'Link' | 'File' | 'Workflow' | 'History' | 'UI'

export interface PreviewReport {
  /** how many objects the action would change */
  readonly objects: number
  readonly files: readonly string[]
  readonly summary: readonly string[]
  readonly warnings: readonly string[]
}

export type ChangeInput = ChangeSet['changes'][number]

export interface ActionContext {
  readonly caller: ActionCaller
  readonly docId: string
  readonly txId: string
  /** records what the action changed; the runner wraps it into a change set */
  emit(changes: readonly ChangeInput[]): void
}

export interface ActionDescriptor<I = unknown, O = unknown> {
  /** `Group.Verb`, e.g. `Data.SetField` */
  readonly name: string
  readonly group: ActionGroup
  readonly summary: string
  readonly input: ZodType<I>
  readonly effect: ActionEffect
  preview(input: I, ctx: ActionContext): Promise<PreviewReport> | PreviewReport
  execute(input: I, ctx: ActionContext): Promise<O> | O
}

export class ActionError extends Error {
  constructor(
    readonly code:
      'unknown_action' | 'invalid_input' | 'permission_denied' | 'confirmation_required',
    message: string,
  ) {
    super(message)
    this.name = 'ActionError'
  }
}

export interface RunOptions {
  readonly caller: ActionCaller
  readonly docId: string
  /** effects the caller may perform (AI and scripts default to read + write) */
  readonly permissions: ReadonlySet<ActionEffect>
  /** validate and preview only */
  readonly dryRun?: boolean
  /** asked before destructive, external or bulk actions; false (or absent) refuses them */
  readonly confirm?: (name: string, preview: PreviewReport) => Promise<boolean> | boolean
  readonly txId?: string
  readonly now?: () => string
}

export interface RunResult<O = unknown> {
  readonly preview: PreviewReport
  readonly output?: O
  readonly changeSet?: ChangeSet
}

const CONFIRM_EFFECTS = new Set<ActionEffect>(['destructive', 'external', 'bulk'])

let sequence = 0
const nextId = (prefix: string) =>
  `${prefix}_${Date.now().toString(36)}${(sequence++).toString(36).padStart(4, '0')}`

export class ActionRegistry {
  private readonly actions = new Map<string, ActionDescriptor<never, unknown>>()

  register<I, O>(action: ActionDescriptor<I, O>): () => void {
    if (this.actions.has(action.name))
      throw new Error(`action ${action.name} is already registered`)
    this.actions.set(action.name, action as unknown as ActionDescriptor<never, unknown>)
    return () => this.actions.delete(action.name)
  }

  has(name: string): boolean {
    return this.actions.has(name)
  }

  list(): ActionDescriptor<unknown, unknown>[] {
    return [...this.actions.values()] as ActionDescriptor<unknown, unknown>[]
  }

  /** The catalog an AI planner or a script editor sees: names, effects and JSON Schema inputs. */
  catalog(): {
    name: string
    group: ActionGroup
    summary: string
    effect: ActionEffect
    input: unknown
  }[] {
    return this.list().map((a) => ({
      name: a.name,
      group: a.group,
      summary: a.summary,
      effect: a.effect,
      input: z.toJSONSchema(a.input as ZodType),
    }))
  }

  /** Validates, checks permission, previews, asks for confirmation when needed, then executes. */
  async run<O = unknown>(
    name: string,
    rawInput: unknown,
    options: RunOptions,
  ): Promise<RunResult<O>> {
    const action = this.actions.get(name) as ActionDescriptor<unknown, O> | undefined
    if (!action) throw new ActionError('unknown_action', `no action named ${name}`)
    const parsed = action.input.safeParse(rawInput)
    if (!parsed.success) throw new ActionError('invalid_input', `${name}: ${parsed.error.message}`)
    if (!options.permissions.has(action.effect)) {
      throw new ActionError(
        'permission_denied',
        `${options.caller} may not run ${action.effect} action ${name}`,
      )
    }
    const changes: ChangeInput[] = []
    const ctx: ActionContext = {
      caller: options.caller,
      docId: options.docId,
      txId: options.txId ?? nextId('tx'),
      emit: (list) => changes.push(...list),
    }
    const preview = await action.preview(parsed.data, ctx)
    if (options.dryRun) return { preview }
    if (CONFIRM_EFFECTS.has(action.effect) && !(await options.confirm?.(name, preview))) {
      throw new ActionError('confirmation_required', `${name} needs confirmation`)
    }
    const output = await action.execute(parsed.data, ctx)
    if (changes.length === 0) return { preview, output }
    const changeSet: ChangeSet = {
      id: nextId('cs'),
      txId: ctx.txId,
      docId: options.docId,
      at: (options.now ?? (() => new Date().toISOString()))(),
      source: options.caller === 'extension' ? 'ui' : options.caller,
      action: name,
      changes,
    }
    return { preview, output, changeSet }
  }
}

/** The default permission sets per caller (AI and scripts never get destructive/external by default). */
export function defaultPermissions(caller: ActionCaller): ReadonlySet<ActionEffect> {
  return caller === 'ui'
    ? new Set<ActionEffect>(['read', 'write', 'bulk', 'destructive', 'external'])
    : new Set<ActionEffect>(['read', 'write'])
}

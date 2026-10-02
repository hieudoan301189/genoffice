/**
 * DVH Action Core (ADR D6, D7): one action catalog shared by the UI, AI,
 * the recorder and DVH-Script. An action validates its input, declares its
 * effect class (the permission it needs), previews without side effects, and
 * reports what it changed as change sets of one transaction.
 */

import { z, type ZodType } from 'zod'
import { contentHash, type ChangeSet } from '@genoffice/dvh-model'

/** the schema builder action inputs are written with, so apps need no zod dependency of their own */
export { z }

export type ActionEffect = 'read' | 'write' | 'bulk' | 'destructive' | 'external'
export type ActionCaller = 'ui' | 'ai' | 'workflow' | 'script' | 'extension'
export type ActionGroup =
  | 'Data'
  | 'Table'
  | 'Document'
  | 'Spreadsheet'
  | 'Presentation'
  | 'Link'
  | 'File'
  | 'Workflow'
  | 'History'
  | 'UI'

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

export type ActionErrorCode =
  | 'unknown_action'
  | 'invalid_input'
  | 'permission_denied'
  | 'confirmation_required'
  | 'stale_catalog'

export class ActionError extends Error {
  constructor(
    readonly code: ActionErrorCode,
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
  /** when to ask (default: destructive and external always, anything over 50 objects) */
  readonly policy?: ConfirmPolicy
  /**
   * the catalog fingerprint the caller planned against (an AI plan, a saved
   * workflow): a different catalog refuses the run instead of guessing
   */
  readonly fingerprint?: string
  readonly txId?: string
  readonly now?: () => string
}

/**
 * Confirmation policy (P4): destructive and external actions are always
 * confirmed; any action whose preview touches more than `bulkThreshold`
 * objects is too, whatever its declared effect. The threshold is a user setting.
 */
export interface ConfirmPolicy {
  readonly bulkThreshold: number
}

export const DEFAULT_CONFIRM_POLICY: ConfirmPolicy = { bulkThreshold: 50 }

/** True when the run must be confirmed by the user before it executes. */
export function needsConfirmation(
  effect: ActionEffect,
  preview: PreviewReport,
  policy: ConfirmPolicy = DEFAULT_CONFIRM_POLICY,
): boolean {
  if (effect === 'destructive' || effect === 'external') return true
  return preview.objects > policy.bulkThreshold
}

/**
 * Catalog fingerprint: 12 hex characters over the canonical JSON of the
 * entries, as the CLI's op catalogs do (packages/cli/src/op-catalog.ts), with
 * the synchronous FNV hash of dvh-model so renderers can compute it too.
 */
export function fingerprintOf(value: unknown): string {
  return contentHash(value).slice(0, 12)
}

export interface CatalogEntry {
  readonly name: string
  readonly group: ActionGroup
  readonly summary: string
  readonly effect: ActionEffect
  /** JSON Schema of the input */
  readonly input: unknown
  /** fingerprint of this entry (name, effect, input schema) */
  readonly fingerprint: string
}

export interface RunResult<O = unknown> {
  readonly preview: PreviewReport
  readonly output?: O
  readonly changeSet?: ChangeSet
}

let sequence = 0
const nextId = (prefix: string) =>
  `${prefix}_${Date.now().toString(36)}${(sequence++).toString(36).padStart(4, '0')}`

export interface RegistryOptions {
  /**
   * Wraps every execute: the app makes the run's txId and caller ambient, so
   * the change sets its own modules record (history, links) carry them too —
   * one AI run or workflow is one transaction to undo (P7).
   */
  around?<T>(ctx: ActionContext, run: () => Promise<T>): Promise<T>
}

export class ActionRegistry {
  private readonly actions = new Map<string, ActionDescriptor<never, unknown>>()

  constructor(private readonly options: RegistryOptions = {}) {}

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

  get(name: string): ActionDescriptor<unknown, unknown> | undefined {
    return this.actions.get(name) as ActionDescriptor<unknown, unknown> | undefined
  }

  /** The catalog an AI planner or a script editor sees: names, effects and JSON Schema inputs. */
  catalog(): CatalogEntry[] {
    return this.list().map((a) => {
      const input = jsonSchemaOf(a.input as ZodType)
      return {
        name: a.name,
        group: a.group,
        summary: a.summary,
        effect: a.effect,
        input,
        fingerprint: fingerprintOf({ name: a.name, effect: a.effect, input }),
      }
    })
  }

  /** One fingerprint for the whole catalog: changes when any action, effect or input schema does. */
  fingerprint(): string {
    return fingerprintOf(this.catalog().map((e) => [e.name, e.fingerprint]))
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
    if (options.fingerprint !== undefined && options.fingerprint !== this.fingerprint()) {
      throw new ActionError(
        'stale_catalog',
        `the action catalog changed since the plan was made (${options.fingerprint} → ${this.fingerprint()}); list the actions again`,
      )
    }
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
    if (
      needsConfirmation(action.effect, preview, options.policy) &&
      !(await options.confirm?.(name, preview))
    ) {
      throw new ActionError('confirmation_required', `${name} needs confirmation`)
    }
    const execute = async () => action.execute(parsed.data, ctx)
    const output = this.options.around ? await this.options.around(ctx, execute) : await execute()
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

/** JSON Schema of an input; inputs that cannot be expressed (custom checks) fall back to "any object". */
function jsonSchemaOf(schema: ZodType): unknown {
  try {
    return z.toJSONSchema(schema, { unrepresentable: 'any' })
  } catch {
    return { type: 'object' }
  }
}

export * from './families'
export * from './saga'
export * from './bridge'
export * from './agent-skill'
export * from './ai-plan'

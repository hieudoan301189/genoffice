/**
 * Adapters (P4): the apps already have op systems with their own validation,
 * dry-run and single-undo transactions — Docs `OpDef` / executeOps, Sheets
 * `WorkbookOperation` / applyChangePlan, Slides pptx-ops / runTxn. They are
 * wrapped, not rewritten: every op becomes one catalog action
 * (`Document.SetFont`, `Spreadsheet.SetCells`, `Presentation.SetText`) whose
 * preview is the op system's dry run and whose execute is one transaction.
 */

import type { ZodType } from 'zod'
import { z } from 'zod'
import type {
  ActionContext,
  ActionEffect,
  ActionGroup,
  ActionRegistry,
  ChangeInput,
  PreviewReport,
} from './index'

/** `setFont` → `SetFont`, `insert_rows` → `InsertRows` */
export function pascalCase(op: string): string {
  return op
    .split(/[_\-\s]+/)
    .filter(Boolean)
    .map((part) => part[0]!.toUpperCase() + part.slice(1))
    .join('')
}

/** ops whose name says they remove content: their action is `destructive` */
const DESTRUCTIVE_OP = /^(delete|remove|clear|drop|truncate)/i

export function effectOfOp(op: string): ActionEffect {
  return DESTRUCTIVE_OP.test(op) ? 'destructive' : 'write'
}

export interface OpFamilyEntry {
  /** the op's own name in its system (`setFont`) */
  readonly op: string
  readonly summary?: string
  /** input schema of the op object without its `op` key; absent: any object */
  readonly input?: ZodType
  readonly effect?: ActionEffect
}

export interface OpRunResult {
  /** what the dry run or the transaction reports, one line per change */
  readonly summary: readonly string[]
  /** objects (blocks, cells, elements) touched; drives the confirmation policy */
  readonly objects: number
  readonly warnings?: readonly string[]
  readonly output?: unknown
  /** what changed, for the change set (the op itself when the system reports nothing finer) */
  readonly changes?: readonly ChangeInput[]
}

export interface OpFamily {
  readonly group: ActionGroup
  readonly entries: readonly OpFamilyEntry[]
  /** runs one op object (`{ op, ...input }`) as one transaction, or plans it when dryRun */
  run(op: Record<string, unknown>, dryRun: boolean, ctx: ActionContext): Promise<OpRunResult>
  /** files an op touches (export paths), for the preview */
  files?(op: Record<string, unknown>): string[]
}

/**
 * Registers one catalog action per op. The action name is
 * `<Group>.<PascalOp>`; an op already registered under that name (a richer
 * hand-written action) is left alone. Returns the names registered.
 */
export function registerOpFamily(registry: ActionRegistry, family: OpFamily): string[] {
  const names: string[] = []
  for (const entry of family.entries) {
    const name = `${family.group}.${pascalCase(entry.op)}`
    if (registry.has(name)) continue
    const input = (entry.input ?? z.record(z.string(), z.unknown())) as ZodType<
      Record<string, unknown>
    >
    const opOf = (value: Record<string, unknown>) => ({ ...value, op: entry.op })
    registry.register<Record<string, unknown>, unknown>({
      name,
      group: family.group,
      summary: entry.summary ?? `${entry.op} (${family.group.toLowerCase()} op)`,
      input,
      effect: entry.effect ?? effectOfOp(entry.op),
      preview: async (value, ctx): Promise<PreviewReport> => {
        const op = opOf(value)
        const planned = await family.run(op, true, ctx)
        return {
          objects: planned.objects,
          files: family.files?.(op) ?? [],
          summary: planned.summary,
          warnings: planned.warnings ?? [],
        }
      },
      execute: async (value, ctx) => {
        const op = opOf(value)
        const done = await family.run(op, false, ctx)
        ctx.emit(
          done.changes ?? [{ objectId: ctx.docId, path: entry.op, before: null, after: value }],
        )
        return done.output ?? { summary: done.summary }
      },
    })
    names.push(name)
  }
  return names
}

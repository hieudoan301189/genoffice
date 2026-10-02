/**
 * Catalog-aware checks of a script (P9): what the editor underlines and what
 * an AI-written script must pass before it runs. Parse errors, actions that
 * are not in the catalog, unknown or missing arguments, literal values of the
 * wrong type, and a warning on actions that need the user's confirmation.
 */

import { walkSteps, type Arg, type Workflow } from '@genoffice/dvh-workflow'
import type { ActionEffect } from '@genoffice/dvh-actions'
import { parseScript } from './parse'

export interface Diagnostic {
  /** 1-based */
  readonly line: number
  /** 1-based */
  readonly column: number
  readonly severity: 'error' | 'warning'
  readonly message: string
}

/** What the checks need from a catalog entry (ActionRegistry.catalog()). */
export interface ScriptCatalogEntry {
  readonly name: string
  readonly effect: ActionEffect
  readonly summary?: string
  /** JSON Schema of the input */
  readonly input: unknown
}

interface ObjectSchema {
  properties?: Record<string, { type?: string | string[]; enum?: unknown[]; description?: string }>
  required?: string[]
  additionalProperties?: unknown
}

/** The argument names, required names and simple types of an action's input schema. */
export function argumentsOf(entry: ScriptCatalogEntry): {
  names: string[]
  required: string[]
  open: boolean
  types: Record<string, string[]>
  descriptions: Record<string, string>
} {
  const schema = (entry.input ?? {}) as ObjectSchema
  const properties = schema.properties ?? {}
  const types: Record<string, string[]> = {}
  const descriptions: Record<string, string> = {}
  for (const [name, prop] of Object.entries(properties)) {
    if (prop.type) types[name] = Array.isArray(prop.type) ? prop.type : [prop.type]
    if (prop.description) descriptions[name] = prop.description
  }
  return {
    names: Object.keys(properties),
    required: schema.required ?? [],
    open:
      !schema.properties ||
      (schema.additionalProperties !== false && schema.additionalProperties !== undefined),
    types,
    descriptions,
  }
}

const typeOfValue = (value: unknown): string =>
  value === null
    ? 'null'
    : Array.isArray(value)
      ? 'array'
      : typeof value === 'number'
        ? Number.isInteger(value)
          ? 'integer'
          : 'number'
        : typeof value

function literalFits(arg: Arg, types: string[] | undefined): boolean {
  if (!types || !('value' in arg)) return true
  const actual = typeOfValue(arg.value)
  return types.some((t) => t === actual || (t === 'number' && actual === 'integer'))
}

/** Checks a compiled workflow against a catalog; `lines` maps step paths to source lines. */
export function checkWorkflow(
  workflow: Workflow,
  catalog: readonly ScriptCatalogEntry[],
  lines: Readonly<Record<string, number>> = {},
): Diagnostic[] {
  const byName = new Map(catalog.map((e) => [e.name, e]))
  const out: Diagnostic[] = []
  for (const { path, step } of walkSteps(workflow.steps)) {
    if (step.kind !== 'action') continue
    const line = lines[path] ?? 1
    const at = (message: string, severity: Diagnostic['severity'] = 'error') =>
      out.push({ line, column: 1, severity, message })
    const entry = byName.get(step.action)
    if (!entry) {
      at(`${step.action} is not an action of this document`)
      continue
    }
    const { names, required, open, types } = argumentsOf(entry)
    for (const key of Object.keys(step.args)) {
      if (!open && !names.includes(key))
        at(`${step.action} has no argument ${key} (it takes ${names.join(', ') || 'none'})`)
      else if (!literalFits(step.args[key]!, types[key]))
        at(`${step.action}: ${key} expects ${types[key]!.join(' or ')}`)
    }
    for (const key of required) {
      if (!(key in step.args)) at(`${step.action} needs ${key}:=`)
    }
    if (entry.effect === 'destructive' || entry.effect === 'external')
      at(`${step.action} is ${entry.effect}: it runs only after the user confirms`, 'warning')
  }
  return out
}

/** Parses and checks a script in one go. */
export function checkScript(
  source: string,
  catalog: readonly ScriptCatalogEntry[],
): {
  workflow: Workflow | null
  lines: Readonly<Record<string, number>>
  diagnostics: Diagnostic[]
} {
  const parsed = parseScript(source)
  if (!parsed.ok) {
    return {
      workflow: null,
      lines: {},
      diagnostics: parsed.errors.map((e) => ({
        line: e.line,
        column: e.column,
        severity: 'error' as const,
        message: e.message,
      })),
    }
  }
  return {
    workflow: parsed.workflow,
    lines: parsed.lines,
    diagnostics: checkWorkflow(parsed.workflow, catalog, parsed.lines),
  }
}

export const hasErrors = (diagnostics: readonly Diagnostic[]) =>
  diagnostics.some((d) => d.severity === 'error')

/**
 * The Workflow Model (P8): what the recorder produces, the block editor edits,
 * the engine runs and DVH-Script (P9) compiles to and prints from. Steps call
 * catalog actions; control flow is ForEach, If, Try and Transaction;
 * arguments are literal JSON values or safe expressions (the P6 evaluator).
 *
 * Steps carry no ids: a step is addressed by its path (`2`, `2.body.0`,
 * `3.else.1`), so a workflow printed as a script and parsed back is equal to
 * the original.
 */

import { z } from 'zod'
import { exprError } from '@genoffice/dvh-template'

export type Json = null | boolean | number | string | Json[] | { [key: string]: Json }

export const jsonSchema: z.ZodType<Json> = z.lazy(() =>
  z.union([
    z.null(),
    z.boolean(),
    z.number(),
    z.string(),
    z.array(jsonSchema),
    z.record(z.string(), jsonSchema),
  ]),
)

/** Names of variables, parameters and argument keys. */
export const IDENT = /^[\p{L}_][\p{L}\p{N}_]*$/u
const ident = z.string().regex(IDENT, 'not a valid name')
const expr = z
  .string()
  .min(1)
  .max(2000)
  .refine((source) => exprError(source) === null, {
    error: (issue) => `bad expression: ${exprError(String(issue.input))}`,
  })

/** An argument: a literal JSON value, or an expression evaluated when the step runs. */
export const argSchema = z.union([
  z.object({ value: jsonSchema }).strict(),
  z.object({ expr }).strict(),
])
export type Arg = z.infer<typeof argSchema>

export type Step =
  | {
      kind: 'action'
      action: string
      args: Record<string, Arg>
      /** the document the action runs in (default: the workflow's document) */
      doc?: Arg
      /** stores the action's output in this variable */
      assign?: string
    }
  | { kind: 'set'; variable: string; value: Arg }
  | { kind: 'if'; branches: { when: string; steps: Step[] }[]; else?: Step[] }
  | { kind: 'forEach'; variable: string; items: Arg; where?: string; steps: Step[] }
  | { kind: 'try'; steps: Step[]; catch: Step[]; errorVariable?: string }
  | { kind: 'transaction'; steps: Step[] }
  | { kind: 'log'; message: Arg }

export const stepSchema: z.ZodType<Step> = z.lazy(() =>
  z.discriminatedUnion('kind', [
    z
      .object({
        kind: z.literal('action'),
        action: z.string().regex(/^[A-Z][A-Za-z]*\.[A-Z][A-Za-z0-9]*$/, 'not an action name'),
        args: z.record(ident, argSchema),
        doc: argSchema.optional(),
        assign: ident.optional(),
      })
      .strict(),
    z.object({ kind: z.literal('set'), variable: ident, value: argSchema }).strict(),
    z
      .object({
        kind: z.literal('if'),
        branches: z.array(z.object({ when: expr, steps: z.array(stepSchema) }).strict()).min(1),
        else: z.array(stepSchema).optional(),
      })
      .strict(),
    z
      .object({
        kind: z.literal('forEach'),
        variable: ident,
        items: argSchema,
        where: expr.optional(),
        steps: z.array(stepSchema),
      })
      .strict(),
    z
      .object({
        kind: z.literal('try'),
        steps: z.array(stepSchema),
        catch: z.array(stepSchema),
        errorVariable: ident.optional(),
      })
      .strict(),
    z.object({ kind: z.literal('transaction'), steps: z.array(stepSchema) }).strict(),
    z.object({ kind: z.literal('log'), message: argSchema }).strict(),
  ]),
)

export const paramTypeSchema = z.enum(['text', 'number', 'boolean', 'record', 'list', 'any'])
export type ParamType = z.infer<typeof paramTypeSchema>

export const paramSchema = z
  .object({
    name: ident,
    type: paramTypeSchema,
    default: jsonSchema.optional(),
    description: z.string().optional(),
  })
  .strict()
export type WorkflowParam = z.infer<typeof paramSchema>

export const workflowSchema = z
  .object({
    format: z.literal('dvh-workflow'),
    version: z.literal(1),
    id: z.string().min(1),
    name: z.string(),
    description: z.string().optional(),
    params: z.array(paramSchema),
    steps: z.array(stepSchema),
    /** the action catalog the workflow was recorded or written against */
    catalog: z.string().optional(),
  })
  .strict()
export type Workflow = z.infer<typeof workflowSchema>

let seq = 0
export const newWorkflowId = () =>
  `wf_${Date.now().toString(36)}${(seq++).toString(36).padStart(3, '0')}`

export function emptyWorkflow(name: string, id = newWorkflowId()): Workflow {
  return { format: 'dvh-workflow', version: 1, id, name, params: [], steps: [] }
}

/** The child step lists of a step, with the path segment that names each. */
export function childLists(step: Step): { key: string; steps: Step[] }[] {
  switch (step.kind) {
    case 'if':
      return [
        ...step.branches.map((b, i) => ({ key: `then${i}`, steps: b.steps })),
        ...(step.else ? [{ key: 'else', steps: step.else }] : []),
      ]
    case 'forEach':
    case 'transaction':
      return [{ key: 'body', steps: step.steps }]
    case 'try':
      return [
        { key: 'body', steps: step.steps },
        { key: 'catch', steps: step.catch },
      ]
    default:
      return []
  }
}

/** Every step with its path, depth first in run order. */
export function walkSteps(
  steps: readonly Step[],
  prefix = '',
): { path: string; step: Step; depth: number }[] {
  const out: { path: string; step: Step; depth: number }[] = []
  const visit = (list: readonly Step[], base: string, depth: number) => {
    list.forEach((step, i) => {
      const path = base ? `${base}.${i}` : String(i)
      out.push({ path, step, depth })
      for (const child of childLists(step)) visit(child.steps, `${path}.${child.key}`, depth + 1)
    })
  }
  visit(steps, prefix, 0)
  return out
}

export interface WorkflowIssue {
  readonly path: string
  readonly message: string
}

/**
 * Checks a workflow beyond its shape: parameter names are unique and every
 * action exists in the catalog (when one is given). Whether a name is bound
 * before it is read is left to the run: expressions may read action outputs.
 */
export function workflowIssues(
  workflow: unknown,
  catalog?: { has(name: string): boolean },
): WorkflowIssue[] {
  const parsed = workflowSchema.safeParse(workflow)
  if (!parsed.success)
    return parsed.error.issues.map((i) => ({ path: i.path.join('.'), message: i.message }))
  const issues: WorkflowIssue[] = []
  const names = new Set<string>()
  for (const p of parsed.data.params) {
    if (names.has(p.name)) issues.push({ path: 'params', message: `parameter ${p.name} twice` })
    names.add(p.name)
  }
  if (catalog) {
    for (const { path, step } of walkSteps(parsed.data.steps)) {
      if (step.kind === 'action' && !catalog.has(step.action))
        issues.push({ path, message: `${step.action} is not in the catalog` })
    }
  }
  return issues
}

/** Parses a workflow, throwing one readable error for the first problem. */
export function parseWorkflow(workflow: unknown): Workflow {
  const parsed = workflowSchema.safeParse(workflow)
  if (!parsed.success) {
    const issue = parsed.error.issues[0]!
    throw new Error(`invalid workflow at ${issue.path.join('.') || '(root)'}: ${issue.message}`)
  }
  return parsed.data
}

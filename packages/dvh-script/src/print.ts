/**
 * Workflow Model → DVH-Script (P9). For a workflow in canonical form,
 * `compileScript(printScript(w)).workflow` equals `w`: nothing is lost, no
 * action, argument or parameter. Any workflow can be put in canonical form
 * first (`canonicalWorkflow`), which changes notation only, never meaning.
 */

import type { Arg, Step, Workflow } from '@genoffice/dvh-workflow'
import { doubleQuoted, literalOf, printLiteral } from './lexical'

const INDENT = '    '
const TYPE_NAMES = {
  text: 'Text',
  number: 'Number',
  boolean: 'Boolean',
  record: 'Record',
  list: 'List',
  any: 'Any',
} as const

export function printArg(arg: Arg): string {
  return 'value' in arg ? printLiteral(arg.value) : arg.expr
}

function printSteps(steps: readonly Step[], depth: number, out: string[]): void {
  const pad = INDENT.repeat(depth)
  for (const step of steps) {
    switch (step.kind) {
      case 'action': {
        const args = Object.entries(step.args)
          .map(([key, arg]) => `${key}:=${printArg(arg)}`)
          .join(', ')
        out.push(
          `${pad}${step.assign ? `${step.assign} = ` : ''}${step.action}(${args})${step.doc ? ` On ${printArg(step.doc)}` : ''}`,
        )
        break
      }
      case 'set':
        out.push(`${pad}${step.variable} = ${printArg(step.value)}`)
        break
      case 'log':
        out.push(`${pad}Log ${printArg(step.message)}`)
        break
      case 'if':
        step.branches.forEach((branch, i) => {
          out.push(`${pad}${i === 0 ? 'If' : 'ElseIf'} ${branch.when} Then`)
          printSteps(branch.steps, depth + 1, out)
        })
        if (step.else) {
          out.push(`${pad}Else`)
          printSteps(step.else, depth + 1, out)
        }
        out.push(`${pad}End If`)
        break
      case 'forEach':
        out.push(
          `${pad}For Each ${step.variable} In ${printArg(step.items)}${step.where ? ` Where ${step.where}` : ''}`,
        )
        printSteps(step.steps, depth + 1, out)
        out.push(`${pad}Next`)
        break
      case 'try':
        out.push(`${pad}Try`)
        printSteps(step.steps, depth + 1, out)
        out.push(`${pad}Catch${step.errorVariable ? ` ${step.errorVariable}` : ''}`)
        printSteps(step.catch, depth + 1, out)
        out.push(`${pad}End Try`)
        break
      case 'transaction':
        out.push(`${pad}Transaction`)
        printSteps(step.steps, depth + 1, out)
        out.push(`${pad}End Transaction`)
        break
    }
  }
}

/** The script of a workflow (header lines, a blank line, then the steps). */
export function printScript(input: Workflow): string {
  const workflow = canonicalWorkflow(input)
  const out: string[] = [
    `Workflow ${printLiteral(workflow.name)}`,
    `Id ${printLiteral(workflow.id)}`,
  ]
  if (workflow.description !== undefined)
    out.push(`Description ${printLiteral(workflow.description)}`)
  if (workflow.catalog !== undefined) out.push(`Catalog ${printLiteral(workflow.catalog)}`)
  for (const p of workflow.params) {
    out.push(
      `Param ${p.name} As ${TYPE_NAMES[p.type]}${p.default !== undefined ? ` = ${printLiteral(p.default)}` : ''}${p.description !== undefined ? ` Description ${printLiteral(p.description)}` : ''}`,
    )
  }
  out.push('')
  printSteps(workflow.steps, 0, out)
  return `${out.join('\n')}\n`
}

/** One expression in script notation: trimmed, text in double quotes. */
function canonicalExpr(expr: string): string {
  return doubleQuoted(expr.trim())
}

function canonicalArg(arg: Arg): Arg {
  if ('value' in arg) return { value: structuredClone(arg.value) }
  const expr = canonicalExpr(arg.expr)
  // an expression that is just a literal is that literal
  const literal = literalOf(expr)
  return literal ? { value: literal.value } : { expr }
}

function canonicalSteps(steps: readonly Step[]): Step[] {
  return steps.map((step): Step => {
    switch (step.kind) {
      case 'action':
        return {
          kind: 'action',
          action: step.action,
          args: Object.fromEntries(Object.entries(step.args).map(([k, a]) => [k, canonicalArg(a)])),
          ...(step.doc ? { doc: canonicalArg(step.doc) } : {}),
          ...(step.assign ? { assign: step.assign } : {}),
        }
      case 'set':
        return { kind: 'set', variable: step.variable, value: canonicalArg(step.value) }
      case 'log':
        return { kind: 'log', message: canonicalArg(step.message) }
      case 'if':
        return {
          kind: 'if',
          branches: step.branches.map((b) => ({
            when: canonicalExpr(b.when),
            steps: canonicalSteps(b.steps),
          })),
          ...(step.else ? { else: canonicalSteps(step.else) } : {}),
        }
      case 'forEach':
        return {
          kind: 'forEach',
          variable: step.variable,
          items: canonicalArg(step.items),
          ...(step.where !== undefined ? { where: canonicalExpr(step.where) } : {}),
          steps: canonicalSteps(step.steps),
        }
      case 'try':
        return {
          kind: 'try',
          steps: canonicalSteps(step.steps),
          catch: canonicalSteps(step.catch),
          ...(step.errorVariable ? { errorVariable: step.errorVariable } : {}),
        }
      case 'transaction':
        return { kind: 'transaction', steps: canonicalSteps(step.steps) }
    }
  })
}

/**
 * The workflow in script notation: expressions trimmed and double-quoted,
 * expressions that are plain literals turned into values, keys in a fixed
 * order. Same behaviour, and exactly what parsing the printed script gives.
 */
export function canonicalWorkflow(workflow: Workflow): Workflow {
  return {
    format: 'dvh-workflow',
    version: 1,
    id: workflow.id,
    name: workflow.name,
    ...(workflow.description !== undefined ? { description: workflow.description } : {}),
    params: workflow.params.map((p) => ({
      name: p.name,
      type: p.type,
      ...(p.default !== undefined ? { default: structuredClone(p.default) } : {}),
      ...(p.description !== undefined ? { description: p.description } : {}),
    })),
    steps: canonicalSteps(workflow.steps),
    ...(workflow.catalog !== undefined ? { catalog: workflow.catalog } : {}),
  }
}

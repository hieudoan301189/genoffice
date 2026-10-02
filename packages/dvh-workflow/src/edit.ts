/**
 * Edits of the block editor (P8), as pure functions over step paths: the
 * editor never mutates a workflow in place, so undo is just the previous value.
 */

import { childLists, type Step, type Workflow } from './model'

interface Slot {
  list: Step[]
  index: number
}

/** Clones the workflow and finds the list holding the step at `path` in the clone. */
function locate(workflow: Workflow, path: string): { copy: Workflow; slot: Slot } {
  const copy = structuredClone(workflow)
  const parts = path.split('.')
  let list = copy.steps
  let i = 0
  for (;;) {
    const index = Number(parts[i])
    if (!Number.isInteger(index) || index < 0) throw new Error(`bad step path ${path}`)
    if (i === parts.length - 1) return { copy, slot: { list, index } }
    const step = list[index]
    const key = parts[i + 1]
    const child = step ? childLists(step).find((c) => c.key === key) : undefined
    if (!child) throw new Error(`bad step path ${path}`)
    list = child.steps
    i += 2
  }
}

export function stepAt(workflow: Workflow, path: string): Step | undefined {
  const { slot } = locate(workflow, path)
  return slot.list[slot.index]
}

export function replaceStep(workflow: Workflow, path: string, step: Step): Workflow {
  const { copy, slot } = locate(workflow, path)
  if (!slot.list[slot.index]) throw new Error(`no step at ${path}`)
  slot.list[slot.index] = structuredClone(step)
  return copy
}

export function removeStep(workflow: Workflow, path: string): Workflow {
  const { copy, slot } = locate(workflow, path)
  slot.list.splice(slot.index, 1)
  return copy
}

/** Inserts a step at `path` (the step there and after it move down); `path` may be one past the end. */
export function insertStep(workflow: Workflow, path: string, step: Step): Workflow {
  const { copy, slot } = locate(workflow, path)
  slot.list.splice(Math.min(slot.index, slot.list.length), 0, structuredClone(step))
  return copy
}

/** Moves a step one place up (-1) or down (+1) inside its list. */
export function moveStep(workflow: Workflow, path: string, by: -1 | 1): Workflow {
  const { copy, slot } = locate(workflow, path)
  const to = slot.index + by
  if (to < 0 || to >= slot.list.length) return workflow
  const [step] = slot.list.splice(slot.index, 1)
  slot.list.splice(to, 0, step!)
  return copy
}

/** Puts the step at `path` inside a new Transaction, Try or If block. */
export function wrapStep(
  workflow: Workflow,
  path: string,
  as: 'transaction' | 'try' | { if: string } | { forEach: string; items: string },
): Workflow {
  const { copy, slot } = locate(workflow, path)
  const step = slot.list[slot.index]
  if (!step) throw new Error(`no step at ${path}`)
  slot.list[slot.index] =
    as === 'transaction'
      ? { kind: 'transaction', steps: [step] }
      : as === 'try'
        ? { kind: 'try', steps: [step], catch: [] }
        : 'if' in as
          ? { kind: 'if', branches: [{ when: as.if, steps: [step] }] }
          : { kind: 'forEach', variable: as.forEach, items: { expr: as.items }, steps: [step] }
  return copy
}

/** One line describing a step, for the block editor and the log. */
export function describeStep(step: Step): string {
  const arg = (a: { value: unknown } | { expr: string }) =>
    'expr' in a ? a.expr : JSON.stringify(a.value)
  switch (step.kind) {
    case 'action': {
      const args = Object.entries(step.args)
        .map(([k, v]) => `${k}:=${arg(v)}`)
        .join(', ')
      return `${step.assign ? `${step.assign} = ` : ''}${step.action}(${args})`
    }
    case 'set':
      return `${step.variable} = ${arg(step.value)}`
    case 'log':
      return `Log ${arg(step.message)}`
    case 'if':
      return `If ${step.branches[0]!.when}`
    case 'forEach':
      return `For Each ${step.variable} In ${arg(step.items)}${step.where ? ` Where ${step.where}` : ''}`
    case 'try':
      return 'Try'
    case 'transaction':
      return 'Transaction'
  }
}

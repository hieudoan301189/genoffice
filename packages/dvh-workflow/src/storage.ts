/**
 * Where workflows live (P8): in the Smart Data model of a template (its
 * customXml part, so the workflow travels with the file) or as
 * `<name>.dvhflow.json` files in a project folder.
 */

import type { DvhModel } from '@genoffice/dvh-model'
import { parseWorkflow, workflowSchema, type Workflow } from './model'

export const WORKFLOW_FILE_SUFFIX = '.dvhflow.json'

/** The model with the workflow added or replaced (by id). */
export function storeWorkflow(model: DvhModel, workflow: Workflow): DvhModel {
  const checked = parseWorkflow(workflow)
  const list = [...(model.workflows ?? [])]
  const at = list.findIndex((w) => w.id === checked.id)
  if (at >= 0) list[at] = checked
  else list.push(checked)
  return { ...model, workflows: list }
}

/** The model without the workflow; `workflows` disappears with the last one. */
export function removeWorkflow(model: DvhModel, id: string): DvhModel {
  const list = (model.workflows ?? []).filter((w) => w.id !== id)
  const { workflows: _drop, ...rest } = model
  return list.length > 0 ? { ...rest, workflows: list } : rest
}

/** Workflows of a model; one written in a format this version cannot read is reported, not fatal. */
export function loadWorkflows(model: DvhModel): { workflows: Workflow[]; skipped: string[] } {
  const workflows: Workflow[] = []
  const skipped: string[] = []
  for (const stored of model.workflows ?? []) {
    const parsed = workflowSchema.safeParse(stored)
    if (parsed.success) workflows.push(parsed.data)
    else skipped.push(stored.name || stored.id)
  }
  return { workflows, skipped }
}

/** The file text of a workflow in a project folder (stable key order, trailing newline). */
export function workflowToJson(workflow: Workflow): string {
  return `${JSON.stringify(parseWorkflow(workflow), null, 2)}\n`
}

export function workflowFromJson(text: string): Workflow {
  return parseWorkflow(JSON.parse(text))
}

/** A file name for a workflow: its name, made safe for Windows, plus the suffix. */
export function workflowFileName(workflow: Pick<Workflow, 'name' | 'id'>): string {
  const base =
    workflow.name
      .replace(/[<>:"/\\|?*]/g, ' ')
      .replace(/\p{Cc}/gu, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/[. ]+$/, '') || workflow.id
  return `${base.slice(0, 120)}${WORKFLOW_FILE_SUFFIX}`
}

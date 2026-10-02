import { z } from 'zod'
import {
  needsConfirmation,
  previewSaga,
  remoteParticipant,
  runSaga,
  type ActionEffect,
  type DvhBridgeRequest,
  type DvhCatalogReply,
  type PreviewReport,
  type SagaParticipant,
  type SagaStep,
} from '@genoffice/dvh-actions'
import type { McpToolDefinition } from '../mcp-server'
import type { EditorFamily } from './formats'

/**
 * DVH Action Core over MCP (P4): `dvh_list_actions`, `dvh_preview`,
 * `dvh_execute`. Every app exposes one catalog (Docs `Document.*`, Sheets
 * `Spreadsheet.*`, Slides `Presentation.*`, plus the Smart Data actions);
 * an external agent lists it, previews a plan of steps, then executes the
 * plan as one saga across the open documents — a failing step puts every
 * touched document back to its checkpoint.
 *
 * Safety: an external agent may read and write; destructive and external
 * actions, and plans touching more objects than the user's threshold, run
 * only after the user confirms the preview in a dialog. A plan made against
 * another catalog (fingerprint) is refused. The tool layer stays free of
 * Electron: tab lookup, the bridges and the dialog are injected.
 */

export type DvhFamily = Extract<EditorFamily, 'docx' | 'xlsx' | 'pptx'>

export interface DvhActionsControl {
  /** resolves a tab id or path to an open document: its webContents id and family */
  resolve(document: string): Promise<{ wcId: number; family: DvhFamily; title: string }>
  /** sends one bridge request to that document's action endpoint */
  send(wcId: number, family: DvhFamily, request: DvhBridgeRequest): Promise<unknown>
  /** shows the plan to the user; true when they accept it */
  confirm(title: string, lines: readonly string[]): Promise<boolean>
  /** the user's threshold: more objects than this need confirmation */
  bulkThreshold(): number
}

/** what an external agent may do without asking */
const BASE_GRANT: ActionEffect[] = ['read', 'write', 'bulk']
const CONFIRMED_GRANT: ActionEffect[] = ['read', 'write', 'bulk', 'destructive', 'external']
const MAX_STEPS = 200

const stepSchema = z.object({
  document: z.string().min(1).describe('tab id or path of an open document (from open_documents)'),
  action: z.string().min(1).describe('action name from dvh_list_actions, e.g. Data.SetField'),
  input: z.record(z.string(), z.unknown()).optional().describe("the action's input object"),
})
type StepArg = z.infer<typeof stepSchema>

interface ResolvedStep {
  readonly step: SagaStep
  readonly title: string
}

interface ResolvedPlan {
  readonly steps: ResolvedStep[]
  /** per docId: how to reach it, and its catalog */
  readonly documents: Map<
    string,
    { wcId: number; family: DvhFamily; title: string; catalog: DvhCatalogReply }
  >
}

async function resolvePlan(control: DvhActionsControl, args: StepArg[]): Promise<ResolvedPlan> {
  if (args.length === 0) throw new Error('steps must not be empty')
  if (args.length > MAX_STEPS) throw new Error(`at most ${MAX_STEPS} steps per plan`)
  const byDocument = new Map<string, string>()
  const documents: ResolvedPlan['documents'] = new Map()
  const steps: ResolvedStep[] = []
  for (const arg of args) {
    let docId = byDocument.get(arg.document)
    if (!docId) {
      const target = await control.resolve(arg.document)
      const catalog = (await control.send(target.wcId, target.family, {
        verb: 'catalog',
      })) as DvhCatalogReply
      docId = catalog.docId
      byDocument.set(arg.document, docId)
      documents.set(docId, { ...target, catalog })
    }
    const doc = documents.get(docId)!
    // refused here, before anything runs: an action the document does not offer
    if (!doc.catalog.actions.some((a) => a.name === arg.action)) {
      throw new Error(
        `"${doc.title}" has no action ${arg.action}; call dvh_list_actions for its catalog`,
      )
    }
    steps.push({ step: { docId, action: arg.action, input: arg.input ?? {} }, title: doc.title })
  }
  return { steps, documents }
}

function participants(
  control: DvhActionsControl,
  plan: ResolvedPlan,
  grant: { allow: ActionEffect[]; confirmed: boolean; fingerprints?: Record<string, string> },
): Map<string, SagaParticipant> {
  const out = new Map<string, SagaParticipant>()
  for (const [docId, doc] of plan.documents) {
    const fingerprint = grant.fingerprints?.[docId]
    out.set(
      docId,
      remoteParticipant({
        docId,
        label: doc.title,
        send: (request) => control.send(doc.wcId, doc.family, request),
        grant: {
          allow: grant.allow,
          confirmed: grant.confirmed,
          bulkThreshold: control.bulkThreshold(),
          ...(fingerprint ? { fingerprint } : {}),
        },
      }),
    )
  }
  return out
}

/** One line per step for the confirmation dialog. */
function planLines(
  plan: ResolvedPlan,
  previews: { preview?: PreviewReport; error?: string }[],
): string[] {
  return plan.steps.map(({ step, title }, i) => {
    const preview = previews[i]?.preview
    const detail = preview
      ? `${preview.objects} object(s)${preview.summary.length ? `: ${preview.summary.slice(0, 3).join('; ')}` : ''}${preview.warnings.length ? ` ⚠ ${preview.warnings.join('; ')}` : ''}`
      : (previews[i]?.error ?? '')
    return `${i + 1}. ${title} — ${step.action}: ${detail}`
  })
}

export function createDvhTools(control: DvhActionsControl | undefined): McpToolDefinition[] {
  if (!control) return []
  const preview = async (plan: ResolvedPlan) =>
    previewSaga(
      plan.steps.map((s) => s.step),
      participants(control, plan, { allow: CONFIRMED_GRANT, confirmed: false }),
      { caller: 'extension', permissions: new Set(CONFIRMED_GRANT) },
    )

  return [
    {
      name: 'dvh_list_actions',
      description:
        'List the DVH actions an open document offers (Smart Data, tables, links and every editing op of ' +
        'its app as Document.* / Spreadsheet.* / Presentation.*), with effect class, JSON Schema input and ' +
        'the catalog fingerprint to pass back to dvh_execute.',
      inputSchema: {
        document: z
          .string()
          .min(1)
          .describe('tab id or path of an open document (from open_documents)'),
        group: z.string().optional().describe('only this group, e.g. Data, Table, Document'),
      },
      handler: async (args) => {
        const target = await control.resolve(String(args.document))
        const catalog = (await control.send(target.wcId, target.family, {
          verb: 'catalog',
        })) as DvhCatalogReply
        const group = typeof args.group === 'string' ? args.group : undefined
        return {
          document: target.title,
          docId: catalog.docId,
          fingerprint: catalog.fingerprint,
          confirmThreshold: control.bulkThreshold(),
          actions: catalog.actions.filter((a) => !group || a.group === group),
        }
      },
    },
    {
      name: 'dvh_preview',
      description:
        'Preview a plan of DVH actions without changing anything: each step is validated and dry-run in its ' +
        'document; reports objects touched, a summary and warnings per step, and whether dvh_execute will ' +
        'ask the user to confirm.',
      inputSchema: { steps: z.array(stepSchema).min(1).max(MAX_STEPS) },
      handler: async (args) => {
        const plan = await resolvePlan(control, args.steps as StepArg[])
        const previews = await preview(plan)
        const confirm = plan.steps.some((s, i) => {
          const entry = plan.documents
            .get(s.step.docId)!
            .catalog.actions.find((a) => a.name === s.step.action)!
          const report = previews[i]?.preview
          return report
            ? needsConfirmation(entry.effect, report, { bulkThreshold: control.bulkThreshold() })
            : false
        })
        return {
          steps: previews.map((p, i) => ({
            document: plan.steps[i]!.title,
            action: p.step.action,
            ...(p.preview ? { preview: p.preview } : {}),
            ...(p.error ? { error: p.error } : {}),
          })),
          needsConfirmation: confirm,
        }
      },
    },
    {
      name: 'dvh_execute',
      description:
        'Execute a plan of DVH actions as one transaction across the open documents: each document is ' +
        'checkpointed before its first step and every touched document is rolled back if any step fails. ' +
        'Destructive/external actions, or steps touching more objects than the user threshold, are shown to ' +
        'the user for confirmation first; a declined plan changes nothing. Pass the fingerprints from ' +
        'dvh_list_actions to refuse running against a changed catalog.',
      inputSchema: {
        steps: z.array(stepSchema).min(1).max(MAX_STEPS),
        fingerprints: z
          .record(z.string(), z.string())
          .optional()
          .describe('docId → catalog fingerprint the plan was made against'),
      },
      handler: async (args) => {
        const plan = await resolvePlan(control, args.steps as StepArg[])
        const fingerprints = (args.fingerprints ?? undefined) as Record<string, string> | undefined
        for (const [docId, expected] of Object.entries(fingerprints ?? {})) {
          const doc = plan.documents.get(docId)
          if (doc && doc.catalog.fingerprint !== expected) {
            throw new Error(
              `the catalog of "${doc.title}" changed (${expected} → ${doc.catalog.fingerprint}); list the actions again`,
            )
          }
        }
        const previews = await preview(plan)
        const failed = previews.findIndex((p) => p.error)
        if (failed >= 0) {
          throw new Error(
            `step ${failed + 1} (${plan.steps[failed]!.step.action}) would fail: ${previews[failed]!.error}; nothing was changed`,
          )
        }
        const threshold = control.bulkThreshold()
        const total = previews.reduce((n, p) => n + (p.preview?.objects ?? 0), 0)
        const mustConfirm =
          total > threshold ||
          plan.steps.some((s, i) => {
            const entry = plan.documents
              .get(s.step.docId)!
              .catalog.actions.find((a) => a.name === s.step.action)!
            return needsConfirmation(entry.effect, previews[i]!.preview!, {
              bulkThreshold: threshold,
            })
          })
        let confirmed = false
        if (mustConfirm) {
          confirmed = await control.confirm('DVH actions', planLines(plan, previews))
          if (!confirmed) throw new Error('the user declined the plan; nothing was changed')
        }
        const result = await runSaga(
          plan.steps.map((s) => s.step),
          participants(control, plan, {
            allow: confirmed ? CONFIRMED_GRANT : BASE_GRANT,
            confirmed,
            ...(fingerprints ? { fingerprints } : {}),
          }),
          { caller: 'extension', permissions: new Set(confirmed ? CONFIRMED_GRANT : BASE_GRANT) },
        )
        return {
          ok: result.ok,
          txId: result.txId,
          ...(result.ok
            ? {}
            : {
                failedAt: (result.failedAt ?? 0) + 1,
                error: result.error,
                rolledBack: result.rolledBack.map((id) => plan.documents.get(id)?.title ?? id),
              }),
          changeSets: result.changeSets.length,
          journal: result.journal,
        }
      },
    },
  ]
}

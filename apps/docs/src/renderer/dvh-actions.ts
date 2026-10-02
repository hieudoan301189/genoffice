/**
 * The first DVH actions in Docs (P1; the unified Action Core is P4). They sit
 * on the same primitives as the Smart Data panel and the apply_ops ops, so a
 * field set by the panel, the agent, MCP or a future script records the same
 * change set. The document's history part is written by dvh-smart-data; the
 * change set a run returns is the caller's receipt of the same changes.
 */
import type { Editor } from '@tiptap/core'
import {
  ActionRegistry,
  defaultPermissions,
  localParticipant,
  registerOpFamily,
  z,
  type ActionContext,
  type PreviewReport,
  type SagaParticipant,
} from '@genoffice/dvh-actions'
import {
  dvhCellStyleSchema,
  dvhColumnSchema,
  fieldText,
  revertPlan,
  scalarSchema,
  type ChangeSet,
  type DvhLink,
  type DvhModel,
} from '@genoffice/dvh-model'
import {
  loadWorkflows,
  parseWorkflow,
  recordsOf,
  removeWorkflow,
  runWorkflow,
  storeWorkflow,
  walkSteps,
  type Json,
  type Workflow,
} from '@genoffice/dvh-workflow'
import { executeOps, opCatalog, type Op } from './ai/ops'
import { allChangeSets, applyHistoryWrite, restoreObject, revertTransaction } from './dvh-history'
import {
  insertColumnControl,
  saveCondition,
  wrapInCondition,
  wrapInRepeat,
} from './dvh-template-designer'
import {
  withDvhTransaction,
  activeDvhDocs,
  ensureModel,
  recordDvhDocsChange,
  fieldOccurrences,
  findField,
  linkSourcePaths,
  sourceFromRead,
  type DvhSource,
} from './dvh-smart-data'
import {
  createDocTable,
  insertDocTable,
  refreshDocTable,
  setDocTableStyle,
  updateLinkFromSource,
} from './dvh-tables'

export interface DocsActionHost {
  editor(): Editor | null
  filePath(): string | null
  /** reads a workbook's model part (window.desktop.dvhReadSource in the app) */
  readSource(path: string): Promise<{ path: string; modelXml: string | null } | null>
  /** the document's current bytes (the template of Document.Generate) */
  buildBytes?(): Promise<Uint8Array | null>
  /** asks the user before a workflow step that needs confirmation (absent: refused) */
  confirm?(name: string, preview: PreviewReport): Promise<boolean> | boolean
}

const preview = (objects: number, summary: string[], warnings: string[] = []): PreviewReport => ({
  objects,
  files: [],
  summary,
  warnings,
})

function requireField(ref: string) {
  const field = findField(activeDvhDocs(), ref)
  if (!field) throw new Error(`unknown field "${ref}"`)
  return field
}

function runOps(host: DocsActionHost, ops: Op[], ctx: ActionContext): string {
  const editor = host.editor()
  if (!editor) throw new Error('no document is open')
  const outcome = executeOps(editor, ops, { source: ctx.caller === 'ai' ? 'ai' : 'ui' })
  if (!outcome.ok) throw new Error(outcome.error ?? 'the edit was rejected')
  return outcome.summary
}

async function readLinked(host: DocsActionHost, link: DvhLink): Promise<DvhSource | null> {
  for (const path of linkSourcePaths(link, host.filePath())) {
    try {
      const source = sourceFromRead(await host.readSource(path))
      if (source && source.model.docId === link.source.docId) return source
    } catch {
      // moved or unreadable: try the next candidate
    }
  }
  return null
}

const fieldRef = z.string().min(1).describe('field name such as Project.Name, or its id')

export function createDocsDvhActions(host: DocsActionHost): ActionRegistry {
  // every run's change sets carry its txId and caller (P7: undo an AI run as a whole)
  const registry = new ActionRegistry({
    around: (ctx, run) => withDvhTransaction(ctx.txId, ctx.caller, run),
  })

  registry.register({
    name: 'Data.ListFields',
    group: 'Data',
    summary: 'List the Smart Data fields of the document with their values and uses',
    input: z.object({}).strict(),
    effect: 'read',
    preview: () => preview(0, []),
    execute: () => {
      const editor = host.editor()
      const occurrences = editor ? fieldOccurrences(editor.state.doc) : []
      return (activeDvhDocs()?.model?.fields ?? []).map((f) => ({
        id: f.id,
        name: f.name,
        type: f.type,
        value: f.value,
        uses: occurrences.filter((o) => o.fieldId === f.id).length,
      }))
    },
  })

  registry.register({
    name: 'Data.GetField',
    group: 'Data',
    summary: 'Read one Smart Data field',
    input: z.object({ field: fieldRef }).strict(),
    effect: 'read',
    preview: () => preview(0, []),
    execute: ({ field }) => ({ ...requireField(field) }),
  })

  registry.register({
    name: 'Data.SetField',
    group: 'Data',
    summary: 'Set a Smart Data field; every occurrence in the document follows',
    input: z.object({ field: fieldRef, value: scalarSchema }).strict(),
    effect: 'write',
    preview: ({ field, value }) => {
      const f = requireField(field)
      return preview(1, [`${f.name}: "${fieldText(f.value)}" → "${fieldText(value)}"`])
    },
    execute: ({ field, value }, ctx) => {
      const f = requireField(field)
      const before = f.value
      const summary = runOps(host, [{ op: 'setSmartField', field: f.id, value }], ctx)
      ctx.emit([{ objectId: f.id, path: 'value', before, after: value }])
      return summary
    },
  })

  registry.register({
    name: 'Document.InsertField',
    group: 'Document',
    summary: 'Insert a Smart Data field into one paragraph (at its end unless afterText is given)',
    input: z
      .object({
        field: fieldRef,
        blockIndex: z.number().int().min(0),
        afterText: z.string().min(1).optional(),
        value: scalarSchema.optional(),
      })
      .strict(),
    effect: 'write',
    preview: ({ field, blockIndex }) => preview(1, [`insert ${field} into block ${blockIndex}`]),
    execute: ({ field, blockIndex, afterText, value }, ctx) => {
      const summary = runOps(
        host,
        [
          {
            op: 'insertSmartField',
            target: { blockIndexes: [blockIndex] },
            field,
            ...(afterText !== undefined ? { afterText } : {}),
            ...(value !== undefined ? { value } : {}),
          },
        ],
        ctx,
      )
      const inserted = requireField(field)
      ctx.emit([{ objectId: inserted.id, path: 'occurrence', before: null, after: inserted.name }])
      return summary
    },
  })

  registry.register({
    name: 'Link.Update',
    group: 'Link',
    summary: 'Pull field values from the linked workbooks into the document',
    input: z.object({ linkId: z.string().min(1).optional() }).strict(),
    effect: 'write',
    preview: ({ linkId }) => {
      const links = (activeDvhDocs()?.model?.links ?? []).filter((l) => !linkId || l.id === linkId)
      return preview(
        links.length,
        links.map((l) => `update from ${l.source.path ?? l.source.relPath ?? l.source.docId}`),
      )
    },
    execute: async ({ linkId }, ctx) => {
      const dvh = activeDvhDocs()
      const editor = host.editor()
      if (!dvh?.model || !editor) throw new Error('no document is open')
      const results: {
        linkId: string
        updated: { fields: number; tables: number; editedTables: readonly string[] } | null
      }[] = []
      for (const link of dvh.model.links.filter((l) => !linkId || l.id === linkId)) {
        const source = await readLinked(host, link)
        if (!source) {
          results.push({ linkId: link.id, updated: null })
          continue
        }
        const before = new Map(dvh.model.fields.map((f) => [f.id, f.value]))
        const updated = updateLinkFromSource(editor, dvh, link, source)
        ctx.emit(
          dvh.model.fields
            .filter((f) => before.get(f.id) !== f.value)
            .map((f) => ({
              objectId: f.id,
              path: 'value',
              before: before.get(f.id),
              after: f.value,
            })),
        )
        results.push({ linkId: link.id, updated })
      }
      return results
    },
  })

  registry.register({
    name: 'Document.InsertTable',
    group: 'Document',
    summary:
      'Insert a DVH.Table showing a linked collection, after the block at blockIndex (default: after the caret)',
    input: z
      .object({
        collection: z.string().min(1).describe('collection name or id'),
        name: z.string().min(1).optional(),
        mode: z.enum(['source', 'destination']).optional(),
        afterBlockIndex: z.number().int().min(-1).optional(),
      })
      .strict(),
    effect: 'write',
    preview: ({ collection }) => preview(1, [`insert a table of ${collection}`]),
    execute: ({ collection, name, mode, afterBlockIndex }, ctx) => {
      const dvh = activeDvhDocs()
      const editor = host.editor()
      if (!dvh || !editor) throw new Error('no document is open')
      const source = ctx.caller === 'ai' ? 'ai' : 'ui'
      const table = createDocTable(
        dvh,
        { collectionId: collection, ...(name ? { name } : {}), mode },
        source,
      )
      insertDocTable(editor, dvh, table, { afterBlockIndex, source })
      ctx.emit([{ objectId: table.id, path: '', before: null, after: table.name }])
      return { id: table.id, name: table.name }
    },
  })

  registry.register({
    name: 'Table.Refresh',
    group: 'Table',
    summary: 'Rebuild a table from its collection; hand edits inside it need force',
    input: z.object({ table: z.string().min(1), force: z.boolean().optional() }).strict(),
    effect: 'write',
    preview: ({ table }) => preview(1, [`refresh ${table}`]),
    execute: ({ table, force }, ctx) => {
      const dvh = activeDvhDocs()
      const editor = host.editor()
      if (!dvh || !editor) throw new Error('no document is open')
      const result = refreshDocTable(editor, dvh, table, {
        force,
        source: ctx.caller === 'ai' ? 'ai' : 'ui',
      })
      if (result.needsConfirm) {
        throw new Error(
          `table ${table} was edited by hand; run again with force: true to overwrite`,
        )
      }
      ctx.emit([{ objectId: table, path: 'render', before: null, after: result.refreshed }])
      return result
    },
  })

  registry.register({
    name: 'Table.SetStyle',
    group: 'Table',
    summary: 'Choose source or destination style and the destination header/body styles',
    input: z
      .object({
        table: z.string().min(1),
        mode: z.enum(['source', 'destination']).optional(),
        header: dvhCellStyleSchema.optional(),
        body: dvhCellStyleSchema.optional(),
        bandFill: z
          .string()
          .regex(/^#[0-9A-Fa-f]{6}$/)
          .optional(),
      })
      .strict(),
    effect: 'write',
    preview: ({ table, mode }) => preview(1, [`style ${table}${mode ? ` → ${mode}` : ''}`]),
    execute: ({ table, ...style }, ctx) => {
      const dvh = activeDvhDocs()
      if (!dvh) throw new Error('no document is open')
      const before = dvh.model?.tables.find((tb) => tb.id === table || tb.name === table)?.style
      const updated = setDocTableStyle(dvh, table, style, ctx.caller === 'ai' ? 'ai' : 'ui')
      ctx.emit([{ objectId: updated.id, path: 'style', before, after: updated.style }])
      return updated.style
    },
  })

  // P5: object history — restore one object, undo one transaction, replace a table's data
  registry.register({
    name: 'History.RestoreObject',
    group: 'History',
    summary: 'Restore one field, collection or table to its state right after a change set',
    input: z.object({ changeSet: z.string().min(1), object: z.string().min(1) }).strict(),
    effect: 'write',
    preview: ({ changeSet, object }) => preview(1, [`restore ${object} as of ${changeSet}`]),
    execute: ({ changeSet, object }, ctx) => {
      const dvh = activeDvhDocs()
      const editor = host.editor()
      if (!dvh || !editor) throw new Error('no document is open')
      const results = restoreObject(editor, dvh, changeSet, object)
      ctx.emit(
        results
          .filter((r) => r.applied)
          .map((r) => ({ objectId: r.objectId, path: r.path, before: null, after: changeSet })),
      )
      return results
    },
  })

  registry.register({
    name: 'History.RevertTransaction',
    group: 'History',
    summary:
      'Undo one transaction (an AI run, a workflow) by its txId; later edits of the same objects block it unless force',
    input: z.object({ txId: z.string().min(1), force: z.boolean().optional() }).strict(),
    effect: 'write',
    preview: ({ txId }) => {
      const plan = revertPlan(allChangeSets(activeDvhDocs()), txId)
      return preview(
        plan.writes.length,
        plan.writes.map((w) => `${w.objectId}.${w.path} ← before ${txId}`),
        plan.conflicts.map((c) => `${c.objectId}.${c.path} changed later (${c.changeSet})`),
      )
    },
    execute: ({ txId, force }, ctx) => {
      const dvh = activeDvhDocs()
      const editor = host.editor()
      if (!dvh || !editor) throw new Error('no document is open')
      const outcome = revertTransaction(editor, dvh, txId, { force })
      if (outcome.conflicts.length > 0 && !force) {
        throw new Error(
          `objects changed after ${txId}: ${outcome.conflicts.map((c) => c.objectId).join(', ')}; run again with force: true to revert anyway`,
        )
      }
      ctx.emit(
        outcome.results
          .filter((r) => r.applied)
          .map((r) => ({
            objectId: r.objectId,
            path: r.path,
            before: null,
            after: `revert ${txId}`,
          })),
      )
      return outcome
    },
  })

  registry.register({
    name: 'Table.ReplaceData',
    group: 'Table',
    summary:
      "Replace the rows (and optionally the columns) of a table's data: its collection, or its own embedded rows",
    input: z
      .object({
        table: z.string().min(1).describe('table name or id'),
        rows: z.array(z.array(scalarSchema)),
        columns: z.array(dvhColumnSchema).optional(),
      })
      .strict(),
    effect: 'write',
    preview: ({ table, rows }) => preview(rows.length, [`replace the data of ${table}`]),
    execute: ({ table, rows, columns }, ctx) => {
      const dvh = activeDvhDocs()
      const editor = host.editor()
      const found = dvh?.model?.tables.find((t) => t.id === table || t.name === table)
      if (!dvh || !editor || !found) throw new Error(`unknown table ${table}`)
      const outcome =
        'collectionId' in found.source
          ? applyHistoryWrite(
              editor,
              dvh,
              {
                objectId: found.source.collectionId,
                path: 'data',
                value: {
                  columns:
                    columns ??
                    dvh.model!.collections.find(
                      (c) => 'collectionId' in found.source && c.id === found.source.collectionId,
                    )?.columns ??
                    [],
                  rows,
                },
              },
              'Table.ReplaceData',
            )
          : applyHistoryWrite(
              editor,
              dvh,
              { objectId: found.id, path: 'definition', value: { ...found, source: { rows } } },
              'Table.ReplaceData',
            )
      if (!outcome.applied) throw new Error(outcome.reason ?? 'not replaced')
      ctx.emit([{ objectId: found.id, path: 'data', before: null, after: rows.length }])
      return outcome
    },
  })

  // P6: Smart Templates — conditions, sections, column values, batch generation, files
  const requireDoc = () => {
    const dvh = activeDvhDocs()
    const editor = host.editor()
    if (!dvh || !editor) throw new Error('no document is open')
    return { dvh, editor }
  }
  const selectBlocks = (editor: Editor, from: number, to: number) => {
    const doc = editor.state.doc
    if (from < 0 || to >= doc.childCount || from > to)
      throw new Error('block range out of the document')
    let start = 0
    for (let i = 0; i < from; i++) start += doc.child(i).nodeSize
    let end = start
    for (let i = from; i <= to; i++) end += doc.child(i).nodeSize
    editor.commands.setTextSelection({ from: start + 1, to: end - 1 })
  }
  const desktop = () => (typeof window === 'undefined' ? undefined : window.desktop)

  registry.register({
    name: 'Template.SetCondition',
    group: 'Document',
    summary:
      'Add or change a Smart Template condition (an expression, e.g. NOT row.ok AND row.qty > 0)',
    input: z
      .object({ expr: z.string().min(1), name: z.string().optional(), id: z.string().optional() })
      .strict(),
    effect: 'write',
    preview: ({ expr }) => preview(1, [`condition ${expr}`]),
    execute: ({ expr, name, id }, ctx) => {
      const { dvh } = requireDoc()
      const saved = saveCondition(dvh, expr, name ?? '', id)
      if ('error' in saved) throw new Error(`expression: ${saved.error}`)
      ctx.emit([{ objectId: saved.id, path: 'expr', before: null, after: expr }])
      return saved
    },
  })

  registry.register({
    name: 'Template.WrapSection',
    group: 'Document',
    summary:
      'Make blocks a Smart Template section: shown only if a condition holds, or repeated per record of a collection (a single table repeats its rows)',
    input: z
      .object({
        fromBlock: z.number().int().min(0),
        toBlock: z.number().int().min(0),
        kind: z.enum(['if', 'repeat']),
        condition: z.string().optional().describe('condition id (if) or record filter (repeat)'),
        collection: z.string().optional().describe('collection id or name (repeat)'),
      })
      .strict(),
    effect: 'write',
    preview: ({ fromBlock, toBlock, kind }) =>
      preview(toBlock - fromBlock + 1, [`${kind} section ${fromBlock}–${toBlock}`]),
    execute: ({ fromBlock, toBlock, kind, condition, collection }, ctx) => {
      const { dvh, editor } = requireDoc()
      selectBlocks(editor, fromBlock, toBlock)
      let ok: boolean
      if (kind === 'if') {
        const found = dvh.model?.conditions?.find((c) => c.id === condition)
        if (!found) throw new Error(`unknown condition ${condition ?? ''}`)
        ok = wrapInCondition(editor, found)
      } else {
        const coll = dvh.model?.collections.find(
          (c) => c.id === collection || c.name === collection,
        )
        if (!coll) throw new Error(`unknown collection ${collection ?? ''}`)
        ok = wrapInRepeat(editor, dvh, coll.id, condition)
      }
      if (!ok) throw new Error('these blocks already belong to a content control')
      ctx.emit([
        { objectId: `blocks:${fromBlock}-${toBlock}`, path: 'section', before: null, after: kind },
      ])
      return { ok }
    },
  })

  registry.register({
    name: 'Template.InsertColumn',
    group: 'Document',
    summary: "Insert a record's column value (Smart Template) at the end of a paragraph",
    input: z
      .object({
        collection: z.string().min(1),
        column: z.string().min(1),
        blockIndex: z.number().int().min(0),
      })
      .strict(),
    effect: 'write',
    preview: ({ collection, column }) => preview(1, [`column ${collection}.${column}`]),
    execute: ({ collection, column, blockIndex }, ctx) => {
      const { dvh, editor } = requireDoc()
      const coll = dvh.model?.collections.find((c) => c.id === collection || c.name === collection)
      const col = coll?.columns.find(
        (c) => c.id === column || c.key === column || c.title === column,
      )
      if (!coll || !col) throw new Error(`unknown column ${collection}.${column}`)
      selectBlocks(editor, blockIndex, blockIndex)
      editor.commands.setTextSelection(editor.state.selection.to)
      if (!insertColumnControl(editor, dvh, coll.id, col.id))
        throw new Error('could not insert the column')
      ctx.emit([{ objectId: col.id, path: 'occurrence', before: null, after: blockIndex }])
      return { column: col.id }
    },
  })

  registry.register({
    name: 'Document.Generate',
    group: 'Document',
    summary:
      'Generate documents from this Smart Template: one per record of a collection (filtered), named by a rule, as docx or PDF, into a folder or a zip with an index',
    input: z
      .object({
        collection: z.string().optional(),
        filter: z.string().optional(),
        nameRule: z
          .string()
          .min(1)
          .describe('file name rule with {expression} holes, e.g. BB-{PAD(INDEX,3)}-{row.code}'),
        format: z.enum(['docx', 'pdf']).optional(),
        output: z.enum(['folder', 'zip']).optional(),
        release: z.enum(['none', 'strip', 'all']).optional(),
      })
      .strict(),
    effect: 'external',
    preview: async ({ collection, filter, nameRule }) => {
      const { dvh } = requireDoc()
      const coll = dvh.model?.collections.find((c) => c.id === collection || c.name === collection)
      const plan = await desktop()?.dvhBatchPlan({
        model: dvh.model!,
        ...(coll ? { collectionId: coll.id } : {}),
        ...(filter ? { filter } : {}),
        nameRule,
      })
      if (!plan || 'error' in plan)
        throw new Error(plan && 'error' in plan ? plan.error : 'unavailable here')
      return {
        objects: plan.count,
        files: plan.names,
        summary: [`${plan.count} document(s)`],
        warnings: [...plan.warnings, ...plan.duplicates.map((d) => `duplicate name ${d}`)],
      }
    },
    execute: async ({ collection, filter, nameRule, format, output, release }, ctx) => {
      const { dvh } = requireDoc()
      const bytes = await host.buildBytes?.()
      const api = desktop()
      if (!bytes || !api) throw new Error('generation is unavailable here')
      const coll = dvh.model?.collections.find((c) => c.id === collection || c.name === collection)
      const result = await api.dvhBatchRun(
        {
          model: dvh.model!,
          ...(coll ? { collectionId: coll.id } : {}),
          ...(filter ? { filter } : {}),
          nameRule,
          format: format ?? 'docx',
          output: output ?? 'folder',
          release: release ?? 'none',
          docPath: host.filePath(),
        },
        bytes.slice().buffer as ArrayBuffer,
      )
      if (!result) throw new Error('cancelled')
      if ('error' in result) throw new Error(result.error)
      ctx.emit([
        { objectId: dvh.model!.docId, path: 'generate', before: null, after: result.output },
      ])
      return result
    },
  })

  registry.register({
    name: 'File.ExportPDF',
    group: 'File',
    summary: 'Export a docx file on disk to PDF next to it',
    input: z.object({ path: z.string().min(1) }).strict(),
    effect: 'external',
    preview: ({ path }) => ({
      objects: 1,
      files: [path.replace(/\.docx$/i, '.pdf')],
      summary: [`export ${path}`],
      warnings: [],
    }),
    execute: async ({ path }, ctx) => {
      const out = await desktop()?.dvhExportPdf(path)
      if (!out) throw new Error('the PDF could not be exported')
      ctx.emit([{ objectId: path, path: 'pdf', before: null, after: out }])
      return out
    },
  })

  registry.register({
    name: 'File.Package',
    group: 'File',
    summary: 'Zip files into one package with an index (the user picks where)',
    input: z.object({ paths: z.array(z.string().min(1)).min(1) }).strict(),
    effect: 'external',
    preview: ({ paths }) => ({
      objects: paths.length,
      files: paths,
      summary: [`package ${paths.length} file(s)`],
      warnings: [],
    }),
    execute: async ({ paths }, ctx) => {
      const result = await desktop()?.dvhPackage(paths)
      if (!result) throw new Error('cancelled')
      ctx.emit([{ objectId: result.path, path: 'package', before: null, after: result.count }])
      return result
    },
  })

  // P4 adapter: every op of the editor's op system is a Document.* action. The
  // op system validates it; preview is its dry run; execute is one transaction
  // (one undo item), exactly as apply_ops runs it.
  registerOpFamily(registry, {
    group: 'Document',
    entries: opCatalog().map((def) => ({ op: def.name, summary: def.signature })),
    run: async (op, dryRun, ctx) => {
      const editor = host.editor()
      if (!editor) throw new Error('no document is open')
      const outcome = executeOps(editor, [op], {
        source: ctx.caller === 'ai' ? 'ai' : 'ui',
        dryRun,
      })
      if (!outcome.ok) throw new Error(outcome.error ?? 'the edit was rejected')
      const changed = outcome.results.reduce((n, r) => n + r.changed, 0)
      return {
        summary: dryRun ? (outcome.plan ?? []) : [outcome.summary],
        // a dry run cannot count matches without applying: one op, one object at least
        objects: dryRun ? 1 : changed,
        output: { summary: outcome.summary, results: outcome.results },
      }
    },
  })

  registerWorkflowActions(registry, host)
  return registry
}

const MAX_WORKFLOW_DEPTH = 8
let workflowDepth = 0

/** The workflow stored in the open document by id or name. */
function findWorkflow(ref: string): Workflow {
  const model = activeDvhDocs()?.model
  const found = model
    ? loadWorkflows(model).workflows.find((w) => w.id === ref || w.name === ref)
    : undefined
  if (!found) throw new Error(`no workflow "${ref}" in this document`)
  return found
}

/**
 * Workflows stored with the document (P8): list, save, delete and run them.
 * A run's steps go through this same registry as caller `workflow`, so they
 * share the run's txId, are never recorded, and roll back as a whole when one
 * fails. Started by the user, a workflow may do what the user may (each
 * destructive or external step is still confirmed); started by the agent or a
 * script, it gets their read + write permissions only.
 */
function registerWorkflowActions(registry: ActionRegistry, host: DocsActionHost): void {
  const paramsOf = (workflow: Workflow, params?: Record<string, Json>, collection?: string) => {
    const out: Record<string, Json> = { ...params }
    if (collection) {
      const coll = activeDvhDocs()?.model?.collections.find(
        (c) => c.id === collection || c.name === collection,
      )
      if (!coll) throw new Error(`unknown collection ${collection}`)
      const list = workflow.params.find((p) => p.type === 'list')
      if (!list) throw new Error(`workflow ${workflow.name} has no list parameter for the records`)
      out[list.name] = recordsOf(coll)
    }
    return out
  }
  const runInput = z
    .object({
      workflow: z.string().min(1).describe('workflow id or name'),
      params: z.record(z.string(), z.json()).optional(),
      collection: z
        .string()
        .optional()
        .describe("collection id or name whose records fill the workflow's list parameter"),
    })
    .strict()
  const execute = async (input: z.infer<typeof runInput>, ctx: ActionContext, dryRun: boolean) => {
    const workflow = findWorkflow(input.workflow)
    const docId = ctx.docId
    // a workflow may run another one, but not without end
    if (workflowDepth >= MAX_WORKFLOW_DEPTH)
      throw new Error(
        `workflows nest deeper than ${MAX_WORKFLOW_DEPTH} (${workflow.name} calls itself?)`,
      )
    workflowDepth++
    const result = await runWorkflow(workflow, {
      params: paramsOf(workflow, input.params as Record<string, Json>, input.collection),
      participants: new Map([[docId, docsSagaParticipant(host, registry, docId)]]),
      docId,
      caller: 'workflow',
      permissions: defaultPermissions(ctx.caller === 'ui' ? 'ui' : 'workflow'),
      txId: ctx.txId,
      dryRun,
      ...(host.confirm ? { confirm: host.confirm } : {}),
    }).finally(() => {
      workflowDepth--
    })
    if (!result.ok)
      throw new Error(
        `workflow ${workflow.name} stopped${result.failedAt ? ` at step ${result.failedAt}` : ''}: ${result.error}`,
      )
    return result
  }

  registry.register({
    name: 'Workflow.List',
    group: 'Workflow',
    summary: 'List the workflows stored in this document with their parameters',
    input: z.object({}).strict(),
    effect: 'read',
    preview: () => preview(0, []),
    execute: () => {
      const model = activeDvhDocs()?.model
      return model
        ? loadWorkflows(model).workflows.map((w) => ({
            id: w.id,
            name: w.name,
            params: w.params,
            steps: w.steps.length,
          }))
        : []
    },
  })

  registry.register({
    name: 'Workflow.Run',
    group: 'Workflow',
    summary:
      'Run a stored workflow (one transaction: rolled back as a whole if a step fails); collection fills its list parameter with records',
    input: runInput,
    effect: 'write',
    preview: async (input, ctx) => {
      const result = await execute(input, ctx, true)
      return {
        objects: result.previews.reduce((n, p) => n + p.preview.objects, 0),
        files: result.previews.flatMap((p) => p.preview.files),
        summary: result.previews.map((p) => `${p.action}: ${p.preview.summary.join('; ')}`),
        warnings: result.previews.flatMap((p) => p.preview.warnings),
      }
    },
    execute: async (input, ctx) => {
      const result = await execute(input, ctx, false)
      return { steps: result.steps, actions: result.changeSets.length, log: result.log.length }
    },
  })

  registry.register({
    name: 'Workflow.Save',
    group: 'Workflow',
    summary:
      'Store a workflow (dvh-workflow v1 JSON) in this document, replacing one with the same id',
    input: z.object({ workflow: z.json() }).strict(),
    effect: 'write',
    preview: ({ workflow }) => {
      const parsed = parseWorkflow(workflow)
      return preview(1, [`save workflow ${parsed.name} (${parsed.steps.length} step(s))`])
    },
    execute: ({ workflow }, ctx) => {
      const parsed = parseWorkflow(workflow)
      const missing = [
        ...new Set(
          walkSteps(parsed.steps).flatMap(({ step }) =>
            step.kind === 'action' && !registry.has(step.action) ? [step.action] : [],
          ),
        ),
      ]
      if (missing.length) throw new Error(`unknown action(s): ${missing.join(', ')}`)
      const dvh = activeDvhDocs()
      if (!dvh) throw new Error('no document is open')
      const ready = ensureModel(dvh)
      const before = ready.model.workflows?.find((w) => w.id === parsed.id)?.name ?? null
      ready.model = storeWorkflow(ready.model, parsed)
      recordDvhDocsChange(dvh, 'Workflow.Save', [
        { objectId: parsed.id, path: '', before, after: parsed.name },
      ])
      ctx.emit([{ objectId: parsed.id, path: '', before, after: parsed.name }])
      return { id: parsed.id }
    },
  })

  registry.register({
    name: 'Workflow.Delete',
    group: 'Workflow',
    summary: 'Remove a workflow from this document',
    input: z.object({ workflow: z.string().min(1) }).strict(),
    effect: 'destructive',
    preview: ({ workflow }) => preview(1, [`delete workflow ${findWorkflow(workflow).name}`]),
    execute: ({ workflow }, ctx) => {
      const found = findWorkflow(workflow)
      const dvh = activeDvhDocs()!
      dvh.model = removeWorkflow(dvh.model!, found.id)
      recordDvhDocsChange(dvh, 'Workflow.Delete', [
        { objectId: found.id, path: '', before: found.name, after: null },
      ])
      ctx.emit([{ objectId: found.id, path: '', before: found.name, after: null }])
      return { id: found.id }
    },
  })
}

/** What a saga restores a document to: its content and its Smart Data. */
interface DocsCheckpoint {
  readonly content: unknown
  readonly model: DvhModel | null
  readonly pending: readonly ChangeSet[]
}

/**
 * The document as a saga participant (P4): the checkpoint is the editor's
 * content and the Smart Data model; restoring puts both back as one undoable
 * edit, so the user can still see (and redo) what the failed saga did.
 */
export function docsSagaParticipant(
  host: DocsActionHost,
  registry: ActionRegistry,
  docId: string,
): SagaParticipant {
  return localParticipant({
    docId,
    ...(host.filePath() ? { label: host.filePath()! } : {}),
    registry,
    checkpoint: (): DocsCheckpoint => {
      const editor = host.editor()
      if (!editor) throw new Error('no document is open')
      const dvh = activeDvhDocs()
      return {
        content: editor.getJSON(),
        model: dvh?.model ? structuredClone(dvh.model) : null,
        pending: [...(dvh?.pending ?? [])],
      }
    },
    restore: (checkpoint) => {
      const editor = host.editor()
      if (!editor) throw new Error('no document is open')
      const cp = checkpoint as DocsCheckpoint
      editor.commands.setContent(cp.content as never, { emitUpdate: true })
      const dvh = activeDvhDocs()
      if (dvh) {
        dvh.model = cp.model ? structuredClone(cp.model) : null
        dvh.pending.splice(0, dvh.pending.length, ...cp.pending)
      }
    },
  })
}

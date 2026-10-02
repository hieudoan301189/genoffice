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
  localParticipant,
  registerOpFamily,
  z,
  type ActionContext,
  type PreviewReport,
  type SagaParticipant,
} from '@genoffice/dvh-actions'
import {
  dvhCellStyleSchema,
  fieldText,
  scalarSchema,
  type ChangeSet,
  type DvhLink,
  type DvhModel,
} from '@genoffice/dvh-model'
import { executeOps, opCatalog, type Op } from './ai/ops'
import {
  activeDvhDocs,
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
  const registry = new ActionRegistry()

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

  return registry
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

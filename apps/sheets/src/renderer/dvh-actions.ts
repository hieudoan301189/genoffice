/**
 * The first DVH actions in Sheets (P1; the unified Action Core is P4), on the
 * same primitives as the Smart Data panel. The workbook's history part is
 * written by dvh-smart-data; the change set a run returns is the caller's
 * receipt of the same changes.
 */
import {
  ActionRegistry,
  localParticipant,
  registerOpFamily,
  z,
  type ActionContext,
  type PreviewReport,
  type SagaParticipant,
} from '@genoffice/dvh-actions'
import { dvhCellStyleSchema, fieldText, scalarSchema } from '@genoffice/dvh-model'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { undoStackDepth } from './undo-carry'
import type { UniverRuntime } from './univer-state'
import {
  boundCellValue,
  createField,
  dvhBindings,
  dvhSheetsContext,
  dvhStateOf,
  quoteSheetName,
  setBoundFieldValue,
  splitBindingFormula,
} from './dvh-smart-data'
import {
  createCollection,
  createTable,
  parseRectFormula,
  rectFormula,
  renderTableInSheet,
  setTableStyle,
} from './dvh-tables'

const preview = (objects: number, summary: string[]): PreviewReport => ({
  objects,
  files: [],
  summary,
  warnings: [],
})

function session() {
  const ctx = dvhSheetsContext()
  const runtime = ctx?.getRuntime() ?? null
  const dvh = dvhStateOf(ctx?.getState())
  if (!ctx || !runtime || !dvh) throw new Error('no workbook is open')
  return { ctx, runtime, dvh }
}

function bindingOf(ref: string) {
  const { runtime, dvh } = session()
  const binding = dvhBindings(runtime, dvh).find((b) => b.field.id === ref || b.field.name === ref)
  if (!binding) throw new Error(`unknown field "${ref}"`)
  return binding
}

const callerSource = (ctx: ActionContext) => (ctx.caller === 'ai' ? 'ai' : 'ui')
const fieldRef = z.string().min(1).describe('field name such as Project.Name, or its id')
/** `B5` on the active sheet, or `Sheet1!B5` / `'Dữ liệu'!$B$5` */
const cellRef = z.string().min(2).describe('cell such as B5 or Sheet1!B5')

/** Resolves a cell reference to a sheet name and zero-based row/column. */
export function parseCellRef(
  ref: string,
  activeSheet: string,
): { sheetName: string; row: number; column: number } {
  const qualified = ref.includes('!') ? ref : `${quoteSheetName(activeSheet)}!${ref}`
  const target = splitBindingFormula(qualified)
  if (!target) throw new Error(`not a single cell: ${ref}`)
  const m = /^([A-Z]+)(\d+)$/.exec(target.cell)!
  let column = 0
  for (const ch of m[1]!) column = column * 26 + (ch.charCodeAt(0) - 64)
  return { sheetName: target.sheet, row: Number(m[2]) - 1, column: column - 1 }
}

/** The workbook's op system, as App wires it for the agent and MCP (planFromOps + applyChangePlan). */
export interface SheetsActionHost {
  applyOps(ops: WorkbookOperation[], dryRun: boolean): Promise<unknown>
}

/** Objects a planned or applied batch touches, from the outcome App returns. */
function countOf(outcome: Record<string, unknown>): number {
  const list = (key: string) =>
    Array.isArray(outcome[key]) ? (outcome[key] as unknown[]).length : 0
  const cells =
    typeof outcome.cellChangeCount === 'number' ? outcome.cellChangeCount : list('cellChanges')
  return cells + list('structuralChanges') + list('formatChanges') + list('sheetRenames')
}

function linesOf(outcome: Record<string, unknown>): string[] {
  const out: string[] = []
  for (const key of ['structuralChanges', 'formatChanges', 'sheetRenames']) {
    if (Array.isArray(outcome[key])) out.push(...(outcome[key] as unknown[]).map(String))
  }
  const cells = countOf(outcome) - out.length
  if (cells > 0) out.push(`${cells} cell(s)`)
  if (typeof outcome.message === 'string') out.push(outcome.message)
  return out
}

export function createSheetsDvhActions(host?: SheetsActionHost): ActionRegistry {
  const registry = new ActionRegistry()

  registry.register({
    name: 'Data.ListFields',
    group: 'Data',
    summary: 'List the Smart Data fields of the workbook with their bound cells and live values',
    input: z.object({}).strict(),
    effect: 'read',
    preview: () => preview(0, []),
    execute: () => {
      const { runtime, dvh } = session()
      return dvhBindings(runtime, dvh).map(({ field, formula, broken }) => ({
        id: field.id,
        name: field.name,
        type: field.type,
        cell: broken ? null : formula,
        value: !broken && formula ? (boundCellValue(runtime, formula) ?? field.value) : field.value,
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
    execute: ({ field }) => {
      const { runtime } = session()
      const binding = bindingOf(field)
      const value =
        !binding.broken && binding.formula
          ? (boundCellValue(runtime, binding.formula) ?? binding.field.value)
          : binding.field.value
      return { ...binding.field, value, cell: binding.broken ? null : binding.formula }
    },
  })

  registry.register({
    name: 'Data.SetField',
    group: 'Data',
    summary: 'Set a Smart Data field: its bound cell and the model value',
    input: z.object({ field: fieldRef, value: scalarSchema }).strict(),
    effect: 'write',
    preview: ({ field, value }) => {
      const binding = bindingOf(field)
      return preview(1, [
        `${binding.field.name}${binding.formula ? ` (${binding.formula})` : ''}: "${fieldText(binding.field.value)}" → "${fieldText(value)}"`,
      ])
    },
    execute: ({ field, value }, actx) => {
      const { ctx, runtime, dvh } = session()
      const binding = bindingOf(field)
      const before = binding.field.value
      setBoundFieldValue(runtime, dvh, binding.field.id, value, callerSource(actx))
      ctx.markPending()
      actx.emit([{ objectId: binding.field.id, path: 'value', before, after: value }])
      return { id: binding.field.id, value }
    },
  })

  registry.register({
    name: 'Spreadsheet.BindField',
    group: 'Spreadsheet',
    summary: 'Create a Smart Data field bound to one cell (the binding follows row/column moves)',
    input: z.object({ name: z.string().min(1), cell: cellRef }).strict(),
    effect: 'write',
    preview: ({ name, cell }) => preview(1, [`bind ${name} to ${cell}`]),
    execute: ({ name, cell }, actx) => {
      const { ctx, runtime, dvh } = session()
      const activeSheet =
        runtime.univerAPI.getActiveWorkbook()?.getActiveSheet().getSheetName() ?? ''
      const target = parseCellRef(cell, activeSheet)
      const field = createField(runtime, dvh, { name, ...target, source: callerSource(actx) })
      ctx.markPending()
      actx.emit([
        { objectId: field.id, path: '', before: null, after: { ...field, binding: cell } },
      ])
      return field
    },
  })

  registry.register({
    name: 'Data.CreateCollection',
    group: 'Data',
    summary: 'Make a range whose first row holds the column titles into a collection',
    input: z
      .object({
        name: z.string().min(1),
        range: z.string().min(2).describe('A1:C10 or Sheet1!A1:C10'),
      })
      .strict(),
    effect: 'write',
    preview: ({ name, range }) => preview(1, [`collection ${name} from ${range}`]),
    execute: ({ name, range }, actx) => {
      const { ctx, runtime, dvh } = session()
      const activeSheet =
        runtime.univerAPI.getActiveWorkbook()?.getActiveSheet().getSheetName() ?? ''
      const rect = parseRectFormula(
        range.includes('!') ? range : `${quoteSheetName(activeSheet)}!${range}`,
      )
      if (!rect || rect.rows < 2) throw new Error(`not a range with a title row: ${range}`)
      const collection = createCollection(runtime, dvh, { name, rect, source: callerSource(actx) })
      ctx.markPending()
      actx.emit([{ objectId: collection.id, path: '', before: null, after: { name, range } }])
      return {
        id: collection.id,
        columns: collection.columns.map((c) => c.title),
        rows: collection.rows.length,
      }
    },
  })

  registry.register({
    name: 'Table.Create',
    group: 'Table',
    summary: 'Define a DVH.Table showing every column of a collection',
    input: z
      .object({
        name: z.string().min(1),
        collection: z.string().min(1).describe('collection name or id'),
        mode: z.enum(['source', 'destination']).optional(),
      })
      .strict(),
    effect: 'write',
    preview: ({ name, collection }) => preview(1, [`table ${name} from ${collection}`]),
    execute: ({ name, collection, mode }, actx) => {
      const { ctx, dvh } = session()
      const table = createTable(dvh, {
        name,
        collectionId: collection,
        mode,
        source: callerSource(actx),
      })
      ctx.markPending()
      actx.emit([{ objectId: table.id, path: '', before: null, after: { name, collection } }])
      return { id: table.id, name: table.name }
    },
  })

  registry.register({
    name: 'Table.Render',
    group: 'Table',
    summary: 'Render a table starting at a cell (values and styles in one write)',
    input: z.object({ table: z.string().min(1), cell: cellRef }).strict(),
    effect: 'write',
    preview: ({ table, cell }) => preview(1, [`render ${table} at ${cell}`]),
    execute: ({ table, cell }, actx) => {
      const { ctx, runtime, dvh } = session()
      const activeSheet =
        runtime.univerAPI.getActiveWorkbook()?.getActiveSheet().getSheetName() ?? ''
      const result = renderTableInSheet(runtime, dvh, table, {
        anchor: parseCellRef(cell, activeSheet),
        source: callerSource(actx),
      })
      ctx.markPending()
      actx.emit([
        { objectId: table, path: 'render', before: null, after: rectFormula(result.rect) },
      ])
      return { range: rectFormula(result.rect), rows: result.rows }
    },
  })

  registry.register({
    name: 'Table.Refresh',
    group: 'Table',
    summary: 'Re-render a table from its collection; hand edits inside it need force',
    input: z.object({ table: z.string().min(1), force: z.boolean().optional() }).strict(),
    effect: 'write',
    preview: ({ table }) => preview(1, [`refresh ${table}`]),
    execute: ({ table, force }, actx) => {
      const { ctx, runtime, dvh } = session()
      const result = renderTableInSheet(runtime, dvh, table, { force, source: callerSource(actx) })
      if (result.needsConfirm) {
        throw new Error(
          `table ${table} was edited by hand; run again with force: true to overwrite`,
        )
      }
      ctx.markPending()
      actx.emit([
        { objectId: table, path: 'render', before: null, after: rectFormula(result.rect) },
      ])
      return { range: rectFormula(result.rect), rows: result.rows }
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
    execute: ({ table, ...style }, actx) => {
      const { ctx, dvh } = session()
      const before = dvh.model?.tables.find((tb) => tb.id === table || tb.name === table)?.style
      const updated = setTableStyle(dvh, table, style, callerSource(actx))
      ctx.markPending()
      actx.emit([{ objectId: updated.id, path: 'style', before, after: updated.style }])
      return updated.style
    },
  })

  // P4 adapter: every workbook DSL operation is a Spreadsheet.* action. Its
  // input is the operation's own schema; preview is planFromOps, execute is
  // one applyChangePlan batch (one undo item).
  if (host) {
    registerOpFamily(registry, {
      group: 'Spreadsheet',
      entries: workbookOperationSchema.options.map((option) => {
        const shape = option as unknown as {
          shape: { op: { value: string } }
          omit(mask: { op: true }): z.ZodType
        }
        return { op: shape.shape.op.value, input: shape.omit({ op: true }) }
      }),
      run: async (op, dryRun) => {
        const outcome = (await host.applyOps([op as WorkbookOperation], dryRun)) as Record<
          string,
          unknown
        >
        if (outcome.ok === false) {
          throw new Error(String(outcome.reason ?? outcome.error ?? 'the edit was rejected'))
        }
        return {
          summary: linesOf(outcome),
          objects: Math.max(1, countOf(outcome)),
          output: outcome,
        }
      },
    })
  }

  return registry
}

/**
 * The workbook as a saga participant (P4): the checkpoint is the depth of its
 * undo stack; restoring undoes back down to it, so every batch the saga
 * applied is undone the way the user would undo it.
 */
export function sheetsSagaParticipant(
  runtime: () => UniverRuntime | null,
  registry: ActionRegistry,
  docId: string,
): SagaParticipant {
  return localParticipant({
    docId,
    registry,
    checkpoint: () => undoStackDepth(runtime()),
    restore: async (checkpoint) => {
      const current = runtime()
      if (!current) throw new Error('no workbook is open')
      const target = Number(checkpoint)
      // bounded: an undo that does not shrink the stack stops the loop
      for (let guard = 0; undoStackDepth(current) > target && guard < 500; guard++) {
        const before = undoStackDepth(current)
        await current.univerAPI.undo()
        if (undoStackDepth(current) >= before) break
      }
      if (undoStackDepth(current) > target) throw new Error('the workbook could not be undone')
    },
  })
}

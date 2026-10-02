/**
 * The first DVH actions in Sheets (P1; the unified Action Core is P4), on the
 * same primitives as the Smart Data panel. The workbook's history part is
 * written by dvh-smart-data; the change set a run returns is the caller's
 * receipt of the same changes.
 */
import { ActionRegistry, z, type ActionContext, type PreviewReport } from '@genoffice/dvh-actions'
import { fieldText, scalarSchema } from '@genoffice/dvh-model'
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

export function createSheetsDvhActions(): ActionRegistry {
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

  return registry
}

/**
 * DVH Smart Data in Sheets (P1, ADR D1/D3/D7/D9).
 *
 * - The file's hidden `_dvh.*` binding names are installed into Univer as
 *   system names, so cut/paste and row/column edits move them live and
 *   formulas can use them; the Name Manager never lists them and the
 *   declarative defined-names save skips them.
 * - On save the editor's copy of those names, the model part (field values
 *   synced from their bound cells) and the change-set history are sent as
 *   `dvhState`; the gateway replaces the file's binding names with them.
 */
import {
  emptyModel,
  fieldDefinedName,
  HISTORY_NS,
  isDvhDefinedName,
  MODEL_NS,
  modelHash,
  newDvhId,
  newStoreItemId,
  parseDvhDefinedName,
  parseHistoryXml,
  parseModelXml,
  serializeHistoryXml,
  serializeModelXml,
  type ChangeSet,
  type DvhField,
  type DvhModel,
  type FieldType,
  type HistoryPart,
  type Scalar,
} from '@genoffice/dvh-model'

import { ICommandService } from '@univerjs/core'
import { RemoveColMutation, RemoveRowMutation, RemoveSheetMutation } from '@univerjs/sheets'

import type { WorkbookDvhParts, WorkbookDvhState } from '../shared/desktop-api'
import { syncCollections } from './dvh-tables'
import { t } from './i18n/locale'
import { univerDefinedNames } from './univer-sync'
import { journalSuppression, type LazyWorkbookState, type UniverRuntime } from './univer-state'

export interface DvhSheetState {
  model: DvhModel | null
  modelStoreItemId: string | null
  history: HistoryPart | null
  historyStoreItemId: string | null
  /** change sets made this session, appended to the history part on save */
  pending: ChangeSet[]
  /** field ids whose binding was found broken, so the warning fires once */
  warnedBroken: Set<string>
}

const states = new WeakMap<LazyWorkbookState, DvhSheetState>()

export function dvhStateOf(state: LazyWorkbookState | null | undefined): DvhSheetState | null {
  return state ? (states.get(state) ?? null) : null
}

interface DefinedNameBuilderHost {
  newDefinedNameBuilder(): { load(param: Record<string, unknown>): { build(): unknown } }
  insertDefinedNameBuilder(param: unknown): void
  deleteDefinedName?(name: string): boolean
}

/** Installs (or replaces) a hidden DVH name as a workbook-scoped system name. */
export function installDvhName(runtime: UniverRuntime, name: string, formula: string): void {
  installName(runtime, name, formula)
}

function installName(runtime: UniverRuntime, name: string, formula: string): void {
  const workbook = runtime.univerAPI.getActiveWorkbook() as unknown as DefinedNameBuilderHost | null
  if (!workbook) return
  workbook.insertDefinedNameBuilder(
    workbook
      .newDefinedNameBuilder()
      .load({ name, formulaOrRefString: formula, localSheetId: 'AllDefaultWorkbook' })
      .build(),
  )
}

/** Reads the session file's DVH content and installs its binding names as system names. */
export async function loadDvhSmartData(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  read: (sessionId: string) => Promise<WorkbookDvhParts>,
): Promise<DvhSheetState> {
  const parts = await read(state.file.sessionId)
  let model: DvhModel | null = null
  let history: HistoryPart | null = null
  try {
    model = parts.model ? parseModelXml(parts.model.xml) : null
  } catch (error) {
    console.warn('DVH model part ignored:', error)
  }
  try {
    history = parts.history ? parseHistoryXml(parts.history.xml) : null
  } catch (error) {
    console.warn('DVH history part ignored:', error)
  }
  const dvh: DvhSheetState = {
    model,
    modelStoreItemId: parts.model?.storeItemId ?? null,
    history,
    historyStoreItemId: parts.history?.storeItemId ?? null,
    pending: [],
    warnedBroken: new Set(),
  }
  journalSuppression.active = true
  try {
    for (const { name, formula } of parts.names) {
      try {
        installName(runtime, name, formula)
      } catch (error) {
        console.warn(`DVH binding ${name} could not be installed:`, error)
      }
    }
  } finally {
    journalSuppression.active = false
  }
  states.set(state, dvh)
  return dvh
}

export interface DvhBinding {
  readonly field: DvhField
  /** `Sheet1!$B$5`, or null when the field has no binding name */
  readonly formula: string | null
  readonly broken: boolean
}

function dvhNames(runtime: UniverRuntime | null): { name: string; formula: string }[] {
  return univerDefinedNames(runtime)
    .filter((n) => isDvhDefinedName(n.getName()))
    .map((n) => ({ name: n.getName(), formula: n.getFormulaOrRefString().replace(/^=/, '') }))
}

export function dvhBindings(runtime: UniverRuntime | null, dvh: DvhSheetState): DvhBinding[] {
  const byField = new Map<string, string>()
  for (const { name, formula } of dvhNames(runtime)) {
    const parsed = parseDvhDefinedName(name)
    if (parsed?.kind === 'field') byField.set(parsed.id, formula)
  }
  return (dvh.model?.fields ?? []).map((field) => {
    const formula = byField.get(field.id) ?? null
    return { field, formula, broken: formula === null || formula.includes('#REF!') }
  })
}

/** `Sheet1!$B$5` / `'Dữ liệu'!$B$5` → sheet name + A1 cell, or null. */
export function splitBindingFormula(formula: string): { sheet: string; cell: string } | null {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)$/.exec(formula.trim())
  if (!m) return null
  return { sheet: (m[1] ?? m[2] ?? '').replace(/''/g, "'"), cell: `${m[3]}${m[4]}` }
}

export function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}

/** The live value of a bound cell (undefined when the binding cannot be resolved). */
export function boundCellValue(runtime: UniverRuntime, formula: string): Scalar | undefined {
  const target = splitBindingFormula(formula)
  if (!target) return undefined
  const sheet = runtime.univerAPI.getActiveWorkbook()?.getSheetByName(target.sheet)
  if (!sheet) return undefined
  const value = sheet.getRange(target.cell).getValue() as unknown
  if (value === undefined || value === '') return null
  return typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string'
    ? value
    : String(value)
}

export function inferFieldType(value: Scalar | undefined): FieldType {
  if (typeof value === 'number') return 'number'
  if (typeof value === 'boolean') return 'boolean'
  return 'text'
}

function record(
  dvh: DvhSheetState,
  action: string,
  changes: ChangeSet['changes'],
  source: ChangeSet['source'] = 'ui',
): void {
  if (!dvh.model || changes.length === 0) return
  dvh.pending.push({
    id: newDvhId('r').replace(/^r_/, 'cs_'),
    txId: newDvhId('r').replace(/^r_/, 'tx_'),
    docId: dvh.model.docId,
    at: new Date().toISOString(),
    source,
    action,
    changes,
  })
}

/** Records a change set made by another DVH module (collections, tables). */
export function recordDvhChange(
  dvh: DvhSheetState,
  action: string,
  changes: ChangeSet['changes'],
  source: ChangeSet['source'] = 'ui',
): void {
  record(dvh, action, changes, source)
}

/** Creates a field bound to one cell (`Sheet!$B$5`); returns the new field. */
export function createField(
  runtime: UniverRuntime,
  dvh: DvhSheetState,
  options: {
    name: string
    sheetName: string
    row: number
    column: number
    type?: FieldType
    source?: ChangeSet['source']
  },
): DvhField {
  const name = options.name.trim()
  if (!name) throw new Error('empty field name')
  if (dvh.model?.fields.some((f) => f.name === name)) throw new Error('duplicate field name')
  dvh.model ??= emptyModel(newDvhId('doc'))
  const column = columnLetters(options.column)
  const formula = `${quoteSheetName(options.sheetName)}!$${column}$${options.row + 1}`
  const value = boundCellValue(runtime, formula) ?? null
  const field: DvhField = {
    id: newDvhId('f'),
    name,
    type: options.type ?? inferFieldType(value),
    value,
    access: 'readwrite',
  }
  installName(runtime, fieldDefinedName(field.id), formula)
  dvh.model.fields.push(field)
  record(
    dvh,
    'Spreadsheet.BindField',
    [{ objectId: field.id, path: '', before: null, after: { ...field, binding: formula } }],
    options.source,
  )
  return field
}

export function removeField(runtime: UniverRuntime, dvh: DvhSheetState, fieldId: string): void {
  const field = dvh.model?.fields.find((f) => f.id === fieldId)
  if (!dvh.model || !field) return
  const workbook = runtime.univerAPI.getActiveWorkbook() as unknown as DefinedNameBuilderHost | null
  workbook?.deleteDefinedName?.(fieldDefinedName(fieldId))
  dvh.model.fields = dvh.model.fields.filter((f) => f.id !== fieldId)
  record(dvh, 'Data.RemoveField', [{ objectId: fieldId, path: '', before: field, after: null }])
}

/** Fields whose bound cell was deleted since the last check (each reported once). */
export function newlyBrokenBindings(runtime: UniverRuntime | null, dvh: DvhSheetState): DvhField[] {
  const broken: DvhField[] = []
  for (const binding of dvhBindings(runtime, dvh)) {
    if (!binding.broken || binding.formula === null || dvh.warnedBroken.has(binding.field.id))
      continue
    dvh.warnedBroken.add(binding.field.id)
    broken.push(binding.field)
    record(dvh, 'Spreadsheet.UnbindField', [
      {
        objectId: binding.field.id,
        path: 'binding',
        before: 'deleted cell',
        after: binding.formula,
      },
    ])
  }
  return broken
}

/**
 * Sets a field: writes its bound cell (an ordinary, undoable cell edit) and
 * the model value. A field whose cell was deleted keeps only the model value.
 */
export function setBoundFieldValue(
  runtime: UniverRuntime,
  dvh: DvhSheetState,
  fieldId: string,
  value: Scalar,
  source: ChangeSet['source'] = 'ui',
): DvhField {
  const binding = dvhBindings(runtime, dvh).find((b) => b.field.id === fieldId)
  if (!binding) throw new Error(`unknown field ${fieldId}`)
  const target = !binding.broken && binding.formula ? splitBindingFormula(binding.formula) : null
  if (target) {
    const sheet = runtime.univerAPI.getActiveWorkbook()?.getSheetByName(target.sheet)
    sheet?.getRange(target.cell).setValue(value ?? '')
  }
  const field = binding.field
  if (field.value !== value) {
    const before = field.value
    field.value = value
    record(
      dvh,
      'Data.SetField',
      [{ objectId: field.id, path: 'value', before, after: value }],
      source,
    )
  }
  return field
}

/** Pulls each bound field's value from its cell; records changes. */
export function syncFieldValues(runtime: UniverRuntime, dvh: DvhSheetState): void {
  if (!dvh.model) return
  const changes: ChangeSet['changes'] = []
  for (const binding of dvhBindings(runtime, dvh)) {
    if (binding.broken || binding.formula === null) continue
    const value = boundCellValue(runtime, binding.formula)
    if (value === undefined || value === binding.field.value) continue
    changes.push({
      objectId: binding.field.id,
      path: 'value',
      before: binding.field.value,
      after: value,
    })
    binding.field.value = value
  }
  record(dvh, 'Spreadsheet.CellEdited', changes)
}

/** Whether the model or history changed this session (counts as a pending edit). */
export function dvhPendingCount(dvh: DvhSheetState | null): number {
  return dvh && dvh.pending.length > 0 ? 1 : 0
}

/** The save payload's dvhState: null for workbooks without any DVH content. */
export function collectDvhState(
  runtime: UniverRuntime | null,
  state: LazyWorkbookState,
): WorkbookDvhState | null {
  const dvh = dvhStateOf(state)
  if (!dvh || !runtime) {
    const names = dvhNames(runtime)
    return names.length === 0 ? null : { names, customXmlParts: [] }
  }
  syncFieldValues(runtime, dvh)
  // may grow or shrink a collection range: read the names afterwards
  syncCollections(runtime, dvh)
  const names = dvhNames(runtime)
  const customXmlParts: WorkbookDvhState['customXmlParts'] = []
  if (dvh.model && dvh.pending.length > 0) {
    dvh.modelStoreItemId ??= newStoreItemId()
    dvh.historyStoreItemId ??= newStoreItemId()
    const history: HistoryPart = {
      docId: dvh.model.docId,
      modelHash: modelHash(dvh.model),
      changes: [...(dvh.history?.changes ?? []), ...dvh.pending],
    }
    customXmlParts.push(
      { ns: MODEL_NS, xml: serializeModelXml(dvh.model), storeItemId: dvh.modelStoreItemId },
      { ns: HISTORY_NS, xml: serializeHistoryXml(history), storeItemId: dvh.historyStoreItemId },
    )
  }
  return { names, customXmlParts }
}

function columnLetters(index: number): string {
  let n = index + 1
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

// ---------------- app wiring ----------------

/** What the Smart Data panel and the binding watch need from the App. */
export interface DvhSheetsContext {
  getRuntime(): UniverRuntime | null
  getState(): LazyWorkbookState | null
  /** a DVH change is pending: let Save write it */
  markPending(): void
  notify(message: string): void
  /** false while a streaming open is still filling the sheets */
  isReady?(): boolean
}

let context: DvhSheetsContext | null = null

export function registerDvhSheetsContext(next: DvhSheetsContext): { dispose(): void } {
  context = next
  return {
    dispose() {
      if (context === next) context = null
    },
  }
}

export function dvhSheetsContext(): DvhSheetsContext | null {
  return context
}

const BINDING_BREAKING_MUTATIONS = new Set<string>([
  RemoveRowMutation.id,
  RemoveColMutation.id,
  RemoveSheetMutation.id,
])

/** Warns right after an edit deletes a bound cell, instead of failing the save later. */
export function installDvhBindingWatch(runtime: UniverRuntime): { dispose(): void } {
  const commands = runtime.univer.__getInjector().get(ICommandService)
  const listener = commands.onCommandExecuted((command) => {
    if (!BINDING_BREAKING_MUTATIONS.has(command.id)) return
    // Univer rewrites the names right after the mutation; check on the next tick
    setTimeout(() => {
      const ctx = context
      const dvh = dvhStateOf(ctx?.getState())
      if (!ctx || !dvh) return
      for (const field of newlyBrokenBindings(ctx.getRuntime(), dvh)) {
        ctx.notify(t('dvhUnboundToast', { name: field.name }))
        ctx.markPending()
      }
    }, 0)
  })
  return { dispose: () => listener.dispose() }
}

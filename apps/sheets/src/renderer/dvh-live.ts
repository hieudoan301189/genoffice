/**
 * Live link publisher (P3 automatic links): after edits settle, the workbook
 * publishes its Smart Data as it stands in the cells, so a document linked to
 * it updates without the workbook being saved. Reads only — the model and the
 * history change on save, as before. The first publish happens as soon as the
 * workbook is ready, so the main process knows which tab holds it.
 *
 * Write-back (P3 Read/Write links): a linked document's field edits arrive
 * here and become `Data.SetField` edits — the bound cell changes like any
 * cell edit (undoable, saved with the workbook). A field changed since the
 * document's base is reported as a conflict instead.
 */
import {
  checkFieldWrite,
  contentHash,
  fieldBaseOf,
  fieldText,
  serializeModelXml,
  sheetsSetFieldsRequestSchema,
  type DvhModel,
  type FieldWriteResult,
  type SheetsSetFieldsReply,
} from '@genoffice/dvh-model'
import { ICommandService } from '@univerjs/core'

import {
  boundCellValue,
  dvhBindings,
  dvhSheetsContext,
  dvhStateOf,
  setBoundFieldValue,
  splitBindingFormula,
} from './dvh-smart-data'
import { readCollectionState } from './dvh-tables'
import type { UniverRuntime } from './univer-state'

const SETTLE_MS = 500
/** until the first publish succeeds (the workbook may still be streaming in) */
const FIRST_PUBLISH_RETRY_MS = 1000
const FIRST_PUBLISH_TRIES = 120

/** The model with field values and collections read from the cells just now. */
export function liveModel(runtime: UniverRuntime, model: DvhModel): DvhModel {
  const ctx = dvhSheetsContext()
  const dvh = dvhStateOf(ctx?.getState())
  const next: DvhModel = structuredClone(model)
  if (dvh) {
    const values = new Map(
      dvhBindings(runtime, dvh)
        .filter((b) => !b.broken && b.formula)
        .map((b) => [b.field.id, boundCellValue(runtime, b.formula!)]),
    )
    for (const field of next.fields) {
      const value = values.get(field.id)
      if (value !== undefined) field.value = value
    }
  }
  next.collections = next.collections.map((collection) => {
    const state = readCollectionState(runtime, collection)
    return state ? { ...collection, columns: state.columns, rows: state.rows } : collection
  })
  return next
}

/** Publishes after each burst of edits (mutations), only when the data changed. */
export function installDvhLivePublisher(runtime: UniverRuntime): { dispose(): void } {
  const commands = runtime.univer.__getInjector().get(ICommandService)
  let timer: ReturnType<typeof setTimeout> | undefined
  let lastHash: string | null = null

  /** true once this workbook's data went out (or was already out) */
  const publish = (): boolean => {
    timer = undefined
    const ctx = dvhSheetsContext()
    const state = ctx?.getState()
    const dvh = dvhStateOf(state)
    const model = dvh?.model
    // a streaming open is still filling the sheet: wait for the full data
    if (!ctx || !state || !model || ctx.isReady?.() === false) return false
    if (model.fields.length === 0 && model.collections.length === 0) return false
    const current = ctx.getRuntime()
    if (!current) return false
    try {
      const modelXml = serializeModelXml(liveModel(current, model))
      const hash = contentHash(modelXml)
      if (hash === lastHash) return true
      lastHash = hash
      window.desktopApi?.dvhPublishLive?.({
        docId: model.docId,
        path: state.file.path ?? null,
        modelXml,
      })
      return true
    } catch (error) {
      console.warn('DVH live publish skipped:', error)
      return false
    }
  }

  let tries = 0
  const firstPublish = setInterval(() => {
    if (publish() || ++tries >= FIRST_PUBLISH_TRIES) clearInterval(firstPublish)
  }, FIRST_PUBLISH_RETRY_MS)

  const listener = commands.onCommandExecuted((command) => {
    if (!command.id.startsWith('sheet.mutation.')) return
    clearTimeout(timer)
    timer = setTimeout(publish, SETTLE_MS)
  })
  const offWrites = window.desktopApi?.onDvhSetFields?.((request) => {
    const reply = handleSetFields(runtime, request)
    if (reply) window.desktopApi?.dvhSetFieldsReply?.(reply)
  })
  return {
    dispose() {
      clearTimeout(timer)
      clearInterval(firstPublish)
      listener.dispose()
      offWrites?.()
    },
  }
}

/** True when the bound cell holds a formula: its value is computed, not written. */
function isFormulaCell(runtime: UniverRuntime, formula: string): boolean {
  const target = splitBindingFormula(formula)
  const sheet = target
    ? runtime.univerAPI.getActiveWorkbook()?.getSheetByName(target.sheet)
    : undefined
  return Boolean(sheet?.getRange(target!.cell).getFormula())
}

/** One write-back request from a linked document; null when it is malformed. */
export function handleSetFields(
  fallbackRuntime: UniverRuntime,
  request: unknown,
): SheetsSetFieldsReply | null {
  const parsed = sheetsSetFieldsRequestSchema.safeParse(request)
  if (!parsed.success) return null
  const { requestId, docId, writes } = parsed.data
  const ctx = dvhSheetsContext()
  const state = ctx?.getState()
  const dvh = dvhStateOf(state)
  const runtime = ctx?.getRuntime() ?? fallbackRuntime
  if (!ctx || !dvh?.model || dvh.model.docId !== docId || ctx.isReady?.() === false) {
    return { requestId, handled: false, path: null, results: [] }
  }
  const bindings = dvhBindings(runtime, dvh)
  const results: FieldWriteResult[] = []
  for (const write of writes) {
    const binding = bindings.find((b) => b.field.id === write.fieldId)
    if (!binding) {
      results.push({ fieldId: write.fieldId, status: 'missing' })
      continue
    }
    const live =
      !binding.broken && binding.formula ? boundCellValue(runtime, binding.formula) : undefined
    const current = fieldBaseOf({
      rev: binding.field.rev,
      value: live === undefined ? binding.field.value : live,
    })
    if (checkFieldWrite(write, current) === 'conflict') {
      results.push({ fieldId: write.fieldId, status: 'conflict', current })
      continue
    }
    if (binding.formula && !binding.broken && isFormulaCell(runtime, binding.formula)) {
      results.push({ fieldId: write.fieldId, status: 'formula', current })
      continue
    }
    const revBefore = binding.field.rev
    const field = setBoundFieldValue(runtime, dvh, write.fieldId, write.value, 'link')
    // a write equal to the model value still changes the cell (it held an unsaved edit)
    if (current.text !== fieldText(write.value) && field.rev === revBefore) {
      field.rev = (field.rev ?? 0) + 1
    }
    results.push({ fieldId: write.fieldId, status: 'written', current: fieldBaseOf(field) })
  }
  if (results.some((r) => r.status === 'written')) ctx.markPending()
  return { requestId, handled: true, path: state?.file.path ?? null, results }
}

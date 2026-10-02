/**
 * Live link publisher (P3 automatic links): after edits settle, the workbook
 * publishes its Smart Data as it stands in the cells, so a document linked to
 * it updates without the workbook being saved. Reads only — the model and the
 * history change on save, as before.
 */
import { contentHash, serializeModelXml, type DvhModel } from '@genoffice/dvh-model'
import { ICommandService } from '@univerjs/core'

import { boundCellValue, dvhBindings, dvhSheetsContext, dvhStateOf } from './dvh-smart-data'
import { readCollectionState } from './dvh-tables'
import type { UniverRuntime } from './univer-state'

const SETTLE_MS = 500

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

  const publish = () => {
    timer = undefined
    const ctx = dvhSheetsContext()
    const state = ctx?.getState()
    const dvh = dvhStateOf(state)
    const model = dvh?.model
    // a streaming open is still filling the sheet: wait for the full data
    if (!ctx || !state || !model || ctx.isReady?.() === false) return
    if (model.fields.length === 0 && model.collections.length === 0) return
    const current = ctx.getRuntime()
    if (!current) return
    try {
      const modelXml = serializeModelXml(liveModel(current, model))
      const hash = contentHash(modelXml)
      if (hash === lastHash) return
      lastHash = hash
      window.desktopApi?.dvhPublishLive?.({
        docId: model.docId,
        path: state.file.path ?? null,
        modelXml,
      })
    } catch (error) {
      console.warn('DVH live publish skipped:', error)
    }
  }

  const listener = commands.onCommandExecuted((command) => {
    if (!command.id.startsWith('sheet.mutation.')) return
    clearTimeout(timer)
    timer = setTimeout(publish, SETTLE_MS)
  })
  return {
    dispose() {
      clearTimeout(timer)
      listener.dispose()
    },
  }
}

import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import { showToast } from './toast-bus'
import type { DvhDocsState } from '../dvh-smart-data'
import {
  insertColumnControl,
  mergeImportedData,
  saveCondition,
  wrapInCondition,
  wrapInRepeat,
} from '../dvh-template-designer'

type Release = 'none' | 'strip' | 'all'

/**
 * Smart Template tools of the Smart Data panel (P6): conditions and repeating
 * sections around the selection, column values at the caret, a preview of one
 * record, batch generation (check, then generate), DVH-Tool conversion and the
 * QLCL exchange.
 */
export function DvhTemplateSection({
  editor,
  dvh,
  filePath,
  fileName,
  buildBytes,
  onChange,
}: {
  editor: Editor
  dvh: DvhDocsState | null
  filePath: string | null
  fileName: string
  /** the document's current bytes: the template */
  buildBytes?: () => Promise<Uint8Array | null>
  onChange: () => void
}) {
  const { t } = useI18n()
  const model = dvh?.model ?? null
  const collections = model?.collections ?? []
  const conditions = model?.conditions ?? []
  const [condName, setCondName] = useState('')
  const [condExpr, setCondExpr] = useState('')
  const [condId, setCondId] = useState('')
  const [collectionId, setCollectionId] = useState('')
  const [repeatFilter, setRepeatFilter] = useState('')
  const [columnId, setColumnId] = useState('')
  const [record, setRecord] = useState('1')
  const [batchCollection, setBatchCollection] = useState('')
  const [filter, setFilter] = useState('')
  const [nameRule, setNameRule] = useState('BB-{PAD(INDEX, 3)}')
  const [format, setFormat] = useState<'docx' | 'pdf'>('docx')
  const [output, setOutput] = useState<'folder' | 'zip'>('folder')
  const [release, setRelease] = useState<Release>('none')
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)

  const collection = collections.find((c) => c.id === collectionId) ?? collections[0]
  const request = () => ({
    model: model!,
    ...(batchCollection ? { collectionId: batchCollection } : {}),
    ...(filter.trim() ? { filter } : {}),
    nameRule,
    release,
    docPath: filePath,
    locale: navigator.language,
  })
  const bytes = async () => {
    const data = await buildBytes?.()
    return data ? (data.slice().buffer as ArrayBuffer) : null
  }
  const run = async (job: () => Promise<void>) => {
    if (busy) return
    setBusy(true)
    try {
      await job()
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
      onChange()
    }
  }

  const save = () => {
    if (!dvh) return
    const result = saveCondition(dvh, condExpr, condName, condId || undefined)
    if ('error' in result) setMessage(result.error)
    else {
      setCondId(result.id)
      setMessage('')
    }
    onChange()
  }

  const wrapIf = () => {
    const condition = conditions.find((c) => c.id === condId)
    if (!condition) return
    if (!wrapInCondition(editor, condition)) setMessage(t('dvhWrapRefused'))
    onChange()
  }

  const wrapRepeat = () => {
    if (!dvh || !collection) return
    if (!wrapInRepeat(editor, dvh, collection.id, repeatFilter || undefined))
      setMessage(t('dvhWrapRefused'))
    onChange()
  }

  const insertColumn = () => {
    if (!dvh || !collection) return
    const column = collection.columns.find((c) => c.id === columnId) ?? collection.columns[0]
    if (column) insertColumnControl(editor, dvh, collection.id, column.id)
    onChange()
  }

  const preview = () =>
    run(async () => {
      const data = await bytes()
      if (!data || !model) return
      const result = await window.desktop.dvhTemplatePreview(
        {
          model,
          ...(batchCollection ? { collectionId: batchCollection } : {}),
          ...(filter.trim() ? { filter } : {}),
          release,
          docPath: filePath,
          locale: navigator.language,
          index: Math.max(0, (Number(record) || 1) - 1),
        },
        data,
      )
      setMessage('error' in result ? result.error : result.warnings.join('\n'))
    })

  const check = () =>
    run(async () => {
      if (!model) return
      const plan = await window.desktop.dvhBatchPlan(request())
      if ('error' in plan) {
        setMessage(plan.error)
        return
      }
      setMessage(
        [
          t('dvhBatchPlanned', { count: plan.count, names: plan.names.slice(0, 5).join(', ') }),
          plan.duplicates.length
            ? t('dvhBatchDuplicates', { names: plan.duplicates.join(', ') })
            : '',
          ...plan.warnings.slice(0, 5),
        ]
          .filter(Boolean)
          .join('\n'),
      )
    })

  const generate = () =>
    run(async () => {
      const data = await bytes()
      if (!data || !model) return
      const result = await window.desktop.dvhBatchRun({ ...request(), format, output }, data)
      if (!result) return
      if ('error' in result) {
        setMessage(result.error)
        return
      }
      showToast(t('dvhBatchDone', { count: result.count, path: result.output }))
      setMessage(result.warnings.slice(0, 5).join('\n'))
    })

  const convert = () =>
    run(async () => {
      const data = await bytes()
      if (!data) return
      const result = await window.desktop.dvhConvertTemplate(data, fileName)
      if (result) {
        showToast(
          t('dvhConverted', {
            fields: result.report.fields.length,
            regions: result.report.regions.length,
          }),
        )
        setMessage(result.report.warnings.join('\n'))
      }
    })

  const importQlcl = () =>
    run(async () => {
      if (!dvh) return
      const result = await window.desktop.dvhQlclImport(dvh.model)
      if (!result) return
      mergeImportedData(editor, dvh, result.model)
      showToast(
        t('dvhQlclImported', {
          fields: result.model.fields.length,
          collections: result.model.collections.length,
          name: result.path.split(/[\\/]/).pop() ?? result.path,
        }),
      )
    })

  const exportQlcl = () =>
    run(async () => {
      if (model) await window.desktop.dvhQlclExport(model)
    })

  return (
    <details className="dvh-docs-panel-section" data-section="template">
      <summary>{t('dvhTemplate')}</summary>

      <div className="dvh-docs-panel-subhead">
        <span>{t('dvhConditions')}</span>
      </div>
      <div className="dvh-docs-panel-new">
        <select
          aria-label={t('dvhConditions')}
          value={condId}
          onChange={(event) => {
            const picked = conditions.find((c) => c.id === event.target.value)
            setCondId(event.target.value)
            setCondName(picked?.name ?? '')
            setCondExpr(picked?.expr ?? '')
          }}
        >
          <option value="">+</option>
          {conditions.map((c) => (
            <option key={c.id} value={c.id}>
              {c.name ?? c.expr}
            </option>
          ))}
        </select>
        <input
          value={condName}
          placeholder={t('dvhCondName')}
          aria-label={t('dvhCondName')}
          onChange={(event) => setCondName(event.target.value)}
        />
        <input
          value={condExpr}
          placeholder={t('dvhCondExpr')}
          aria-label={t('dvhCondExpr')}
          onChange={(event) => setCondExpr(event.target.value)}
        />
        <button type="button" disabled={!condExpr.trim()} onClick={save}>
          {t('dvhCondSave')}
        </button>
        <button type="button" disabled={!condId} onClick={wrapIf}>
          {t('dvhWrapIf')}
        </button>
      </div>

      <div className="dvh-docs-panel-new">
        <select
          aria-label={t('dvhWrapRepeat')}
          value={collection?.id ?? ''}
          onChange={(event) => setCollectionId(event.target.value)}
        >
          {collections.map((c) => (
            <option key={c.id} value={c.id}>
              {c.name}
            </option>
          ))}
        </select>
        <select
          aria-label={t('dvhRepeatFilter')}
          value={repeatFilter}
          onChange={(event) => setRepeatFilter(event.target.value)}
        >
          <option value="">{t('dvhAllRecords')}</option>
          {conditions.map((c) => (
            <option key={c.id} value={c.id}>
              {c.name ?? c.expr}
            </option>
          ))}
        </select>
        <button type="button" disabled={!collection} onClick={wrapRepeat}>
          {t('dvhWrapRepeat')}
        </button>
        <select
          aria-label={t('dvhInsertColumn')}
          value={columnId}
          onChange={(event) => setColumnId(event.target.value)}
        >
          {(collection?.columns ?? []).map((c) => (
            <option key={c.id} value={c.id}>
              {c.title}
            </option>
          ))}
        </select>
        <button type="button" disabled={!collection} onClick={insertColumn}>
          {t('dvhInsertColumn')}
        </button>
      </div>

      <div className="dvh-docs-panel-subhead">
        <span>{t('dvhBatch')}</span>
      </div>
      <div className="dvh-docs-panel-new" data-status="batch">
        <select
          aria-label={t('dvhBatchCollection')}
          value={batchCollection}
          onChange={(event) => setBatchCollection(event.target.value)}
        >
          <option value="">{t('dvhBatchSingle')}</option>
          {collections.map((c) => (
            <option key={c.id} value={c.id}>
              {t('dvhBatchCollection')} {c.name}
            </option>
          ))}
        </select>
        <input
          value={filter}
          placeholder={t('dvhBatchFilter')}
          aria-label={t('dvhBatchFilter')}
          onChange={(event) => setFilter(event.target.value)}
        />
        <input
          value={nameRule}
          placeholder={t('dvhBatchName')}
          aria-label={t('dvhBatchName')}
          onChange={(event) => setNameRule(event.target.value)}
        />
        <select
          aria-label={t('dvhBatchFormat')}
          value={format}
          onChange={(event) => setFormat(event.target.value as 'docx' | 'pdf')}
        >
          <option value="docx">docx</option>
          <option value="pdf">PDF</option>
        </select>
        <select
          aria-label={t('dvhBatchOutput')}
          value={output}
          onChange={(event) => setOutput(event.target.value as 'folder' | 'zip')}
        >
          <option value="folder">{t('dvhBatchFolder')}</option>
          <option value="zip">{t('dvhBatchZip')}</option>
        </select>
        <select
          aria-label={t('dvhExportCopy')}
          value={release}
          onChange={(event) => setRelease(event.target.value as Release)}
        >
          <option value="none">{t('dvhBatchKeep')}</option>
          <option value="strip">{t('dvhPolicyStrip')}</option>
          <option value="all">{t('dvhPolicyAll')}</option>
        </select>
        <input
          value={record}
          type="number"
          min={1}
          aria-label={t('dvhPreviewRecord')}
          onChange={(event) => setRecord(event.target.value)}
        />
        <button
          type="button"
          disabled={busy || !model || !buildBytes}
          onClick={() => void preview()}
        >
          {t('dvhPreviewRecord')}
        </button>
        <button type="button" disabled={busy || !model} onClick={() => void check()}>
          {t('dvhBatchCheck')}
        </button>
        <button
          type="button"
          disabled={busy || !model || !buildBytes}
          onClick={() => void generate()}
        >
          {t('dvhBatchRun')}
        </button>
      </div>

      <div className="dvh-docs-panel-link-actions">
        <button type="button" disabled={busy || !buildBytes} onClick={() => void convert()}>
          {t('dvhConvertDvhTool')}
        </button>
        <button type="button" disabled={busy || !dvh} onClick={() => void importQlcl()}>
          {t('dvhQlclImport')}
        </button>
        <button type="button" disabled={busy || !model} onClick={() => void exportQlcl()}>
          {t('dvhQlclExport')}
        </button>
      </div>
      {message ? <p className="dvh-docs-panel-muted dvh-docs-panel-pre">{message}</p> : null}
    </details>
  )
}

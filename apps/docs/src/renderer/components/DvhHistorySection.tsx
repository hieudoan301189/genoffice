import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import {
  cellDiff,
  fieldText,
  type ChangeSet,
  type DvhCollection,
  type Scalar,
} from '@genoffice/dvh-model'
import { useI18n } from '../i18n/locale'
import { showToast } from './toast-bus'
import {
  adoptRecoveredHistory,
  compactDvhHistory,
  discardRecoveredHistory,
  historyEntries,
  historyStatus,
  type DvhDocsState,
} from '../dvh-smart-data'
import { restoreObject, revertTransaction } from '../dvh-history'
import { noteUnrecordable } from '../dvh-workflow'

const HISTORY_ROWS = 30
/** compaction is offered from this share of the limit on */
const NEAR_LIMIT = 0.8

type ReleaseKind = 'all' | 'from' | 'none' | 'strip'

const kb = (bytes: number) =>
  bytes >= 1024 * 1024 ? `${(bytes / 1024 / 1024).toFixed(1)} MB` : `${Math.ceil(bytes / 1024)} KB`

const SOURCE_KEYS = {
  ui: 'dvhSrcUi',
  ai: 'dvhSrcAi',
  workflow: 'dvhSrcWorkflow',
  script: 'dvhSrcScript',
  link: 'dvhSrcLink',
  restore: 'dvhSrcRestore',
  external: 'dvhSrcExternal',
} as const

/**
 * History of the document's Smart Data (P5): per object (field, collection,
 * table, link) or for everything; who/what made each change; value and table
 * diffs; restore one object to any point; undo one run by its txId; size
 * against the limit with compaction; changes a crashed session left behind;
 * and copies for sending out (with all, part or none of the DVH data).
 */
export function DvhHistorySection({
  editor,
  dvh,
  focus,
  onFocus,
  onChange,
  exportCopy,
}: {
  editor: Editor
  dvh: DvhDocsState | null
  /** the object the list is filtered to (null: all) */
  focus: string | null
  onFocus: (objectId: string | null) => void
  onChange: () => void
  /** saves a copy of the document with a release policy; null when cancelled */
  exportCopy?: (
    policy: { kind: 'all' | 'none' | 'strip' } | { kind: 'from'; at: string },
  ) => Promise<string | null>
}) {
  const { t } = useI18n()
  const [release, setRelease] = useState<ReleaseKind>('none')
  const [fromDate, setFromDate] = useState('')
  const model = dvh?.model ?? null
  const objects = [
    ...(model?.fields ?? []).map((f) => ({ id: f.id, name: f.name })),
    ...(model?.collections ?? []).map((c) => ({ id: c.id, name: c.name })),
    ...(model?.tables ?? []).map((tb) => ({ id: tb.id, name: tb.name })),
    ...(model?.links ?? []).map((l) => ({
      id: l.id,
      name: (l.source.path ?? l.source.relPath ?? l.source.docId).split(/[\\/]/).pop()!,
    })),
  ]
  const nameOf = (id: string) => objects.find((o) => o.id === id)?.name ?? id
  const entries = historyEntries(dvh)
    .filter((cs) => !focus || cs.changes.some((c) => c.objectId === focus))
    .slice(0, HISTORY_ROWS)
  const status = historyStatus(dvh)
  const recovered = dvh?.recovered ?? []

  const show = (v: unknown) =>
    v === null || ['string', 'number', 'boolean'].includes(typeof v) ? fieldText(v as Scalar) : '…'

  const describe = (cs: ChangeSet) =>
    cs.changes
      .filter((c) => !focus || c.objectId === focus)
      .slice(0, 6)
      .map((c, i) => {
        const key = `${cs.id}:${i}`
        if (c.path === 'value') {
          return (
            <div key={key} className="dvh-docs-panel-diff">
              {nameOf(c.objectId)}: <del>{show(c.before)}</del> → <ins>{show(c.after)}</ins>
            </div>
          )
        }
        if (c.path === 'data') {
          const diff = cellDiff(
            c.before as Pick<DvhCollection, 'columns' | 'rows'> | null,
            c.after as Pick<DvhCollection, 'columns' | 'rows'> | null,
          )
          const parts = [
            diff.rowsBefore !== diff.rowsAfter
              ? t('dvhDiffRows', { before: diff.rowsBefore, after: diff.rowsAfter })
              : '',
            diff.cells.length ? t('dvhDiffCells', { count: diff.cells.length }) : '',
            diff.columnsAdded.length || diff.columnsRemoved.length
              ? t('dvhDiffColumns', {
                  added: diff.columnsAdded.join(', ') || '0',
                  removed: diff.columnsRemoved.join(', ') || '0',
                })
              : '',
            diff.styleChanged ? t('dvhDiffStyle') : '',
          ].filter(Boolean)
          const sample = diff.cells.slice(0, 3)
          return (
            <div key={key} className="dvh-docs-panel-diff">
              {nameOf(c.objectId)}: {parts.join(', ')}
              {sample.map((cell) => (
                <div key={`${cell.row}:${cell.column}`} className="dvh-docs-panel-muted">
                  {cell.column} #{cell.row + 1}: <del>{show(cell.before)}</del> →{' '}
                  <ins>{show(cell.after)}</ins>
                </div>
              ))}
            </div>
          )
        }
        if (c.path === 'definition') {
          return (
            <div key={key} className="dvh-docs-panel-diff">
              {nameOf(c.objectId)}: {t('dvhDiffDefinition')}
            </div>
          )
        }
        return null
      })

  const restore = (cs: ChangeSet) => {
    if (!dvh || !focus) return
    // a restore names one past change set: replaying it elsewhere means nothing
    noteUnrecordable(t('dvhRestoreHere'))
    try {
      restoreObject(editor, dvh, cs.id, focus)
    } catch (e) {
      showToast(e instanceof Error ? e.message : String(e))
    }
    onChange()
  }

  const revert = (cs: ChangeSet) => {
    if (!dvh) return
    noteUnrecordable(t('dvhRevertTx'))
    try {
      let outcome = revertTransaction(editor, dvh, cs.txId)
      if (outcome.conflicts.length > 0) {
        const names = [...new Set(outcome.conflicts.map((c) => nameOf(c.objectId)))].join(', ')
        if (!window.confirm(t('dvhRevertConflict', { names }))) return
        outcome = revertTransaction(editor, dvh, cs.txId, { force: true })
      }
    } catch (e) {
      showToast(e instanceof Error ? e.message : String(e))
    }
    onChange()
  }

  const runExport = async () => {
    if (!exportCopy) return
    const policy =
      release === 'from'
        ? fromDate
          ? { kind: 'from' as const, at: new Date(fromDate).toISOString() }
          : null
        : { kind: release }
    if (!policy) return
    const path = await exportCopy(policy)
    if (path) showToast(t('dvhExported', { name: path.split(/[\\/]/).pop() ?? path }))
  }

  return (
    <details className="dvh-docs-panel-section dvh-docs-panel-history" open={focus !== null}>
      <summary>{focus ? t('dvhHistoryFor', { name: nameOf(focus) }) : t('dvhHistory')}</summary>

      {recovered.length > 0 && dvh ? (
        <div className="dvh-docs-panel-conflict" data-status="recovered">
          {t('dvhRecovered', { count: recovered.length })}
          <div className="dvh-docs-panel-link-actions">
            <button
              type="button"
              onClick={() => {
                adoptRecoveredHistory(dvh)
                onChange()
              }}
            >
              {t('dvhRecoveredKeep')}
            </button>
            <button
              type="button"
              onClick={() => {
                discardRecoveredHistory(dvh)
                onChange()
              }}
            >
              {t('dvhRecoveredDrop')}
            </button>
          </div>
        </div>
      ) : null}

      <div className="dvh-docs-panel-link-actions">
        <select
          aria-label={t('dvhHistoryObject')}
          value={focus ?? ''}
          onChange={(event) => onFocus(event.target.value || null)}
        >
          <option value="">{t('dvhHistoryAllObjects')}</option>
          {objects.map((o) => (
            <option key={o.id} value={o.id}>
              {o.name}
            </option>
          ))}
        </select>
        <span className="dvh-docs-panel-muted" data-status="history-size">
          {t('dvhHistorySize', { size: kb(status.bytes), limit: kb(status.limit) })}
        </span>
        {dvh && status.bytes >= status.limit * NEAR_LIMIT ? (
          <button
            type="button"
            title={t('dvhHistoryNearLimit')}
            onClick={() => {
              compactDvhHistory(dvh)
              onChange()
            }}
          >
            {t('dvhHistoryCompact')}
          </button>
        ) : null}
      </div>

      {entries.length === 0 ? (
        <p className="dvh-docs-panel-empty">{t('dvhNoHistory')}</p>
      ) : (
        <ul>
          {entries.map((cs) => (
            <li key={cs.id} data-action={cs.action} data-source={cs.source}>
              <span className="dvh-docs-panel-muted">{new Date(cs.at).toLocaleString()}</span>{' '}
              <code>{cs.action}</code>{' '}
              <span className="dvh-docs-panel-badge" data-source={cs.source}>
                {t(SOURCE_KEYS[cs.source])}
              </span>
              {cs.actor?.user ? (
                <span className="dvh-docs-panel-muted"> · {cs.actor.user}</span>
              ) : null}
              {describe(cs)}
              <div className="dvh-docs-panel-link-actions">
                {focus ? (
                  <button type="button" onClick={() => restore(cs)}>
                    {t('dvhRestoreHere')}
                  </button>
                ) : null}
                {cs.source !== 'external' && cs.source !== 'restore' ? (
                  <button type="button" onClick={() => revert(cs)}>
                    {t('dvhRevertTx')}
                  </button>
                ) : null}
              </div>
            </li>
          ))}
        </ul>
      )}

      {exportCopy ? (
        <div className="dvh-docs-panel-new" data-status="release">
          <select
            aria-label={t('dvhExportCopy')}
            value={release}
            onChange={(event) => setRelease(event.target.value as ReleaseKind)}
          >
            <option value="none">{t('dvhPolicyNone')}</option>
            <option value="strip">{t('dvhPolicyStrip')}</option>
            <option value="from">{t('dvhPolicyFrom')}</option>
            <option value="all">{t('dvhPolicyAll')}</option>
          </select>
          {release === 'from' ? (
            <input
              type="date"
              aria-label={t('dvhPolicyFrom')}
              value={fromDate}
              onChange={(event) => setFromDate(event.target.value)}
            />
          ) : null}
          <button type="button" onClick={() => void runExport()}>
            {t('dvhExportCopy')}
          </button>
          <p className="dvh-docs-panel-muted">{t('dvhExportWarn')}</p>
        </div>
      ) : null}
    </details>
  )
}

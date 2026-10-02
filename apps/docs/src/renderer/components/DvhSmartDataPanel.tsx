import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { fieldText, newDvhId, type DvhLink, type Scalar } from '@genoffice/dvh-model'
import { useI18n } from '../i18n/locale'
import { showToast } from './toast-bus'
import { createDocsDvhActions } from '../dvh-actions'
import {
  activeDvhDocs,
  dvhDocsStateOf,
  fieldOccurrences,
  historyEntries,
  insertSmartField,
  linkSource,
  loadDvhDocs,
  relocateLink,
  setActiveDvhDocs,
  setDvhModelChangeListener,
  setLinkUpdateMode,
  sourceFromRead,
  type LinkUpdateMode,
} from '../dvh-smart-data'
import { refreshDocTable, setDvhNumberLocale, unlinkSource } from '../dvh-tables'
import {
  ambiguousSources,
  applySource,
  checkLinksOnOpen,
  installAutoLinks,
  linkConflicts,
  linkPendingWrites,
  linkStatus,
  noteDocumentEdited,
  onLinkStatus,
  refreshAutoWatch,
  resolveConflict,
  resolveLinkedSource,
  setLinkStatus,
  writeBackLink,
  type WriteBackOutcome,
} from '../dvh-auto'
import { DvhDocsTablesSection } from './DvhDocsTablesSection'
import './dvh-smart-data.css'

const fileName = (path: string) => path.split(/[\\/]/).pop() ?? path
const HISTORY_ROWS = 20
const readSource = (path: string) => window.desktop.dvhReadSource({ path })
const findSource = (docId: string) => window.desktop.dvhFindSource?.(docId) ?? Promise.resolve([])

/**
 * Mounted once by App: loads the document's Smart Data whenever a document is
 * (re)parsed, so saves keep the model in step with field text even when the
 * panel was never opened, and toggles the panel on the ribbon's event.
 */
export function DvhSmartDataHost({
  editor,
  parsed,
  filePath,
  markDirty,
}: {
  editor: Editor | null
  parsed: ParsedDocFull | null
  filePath: string | null
  markDirty: () => void
}) {
  const [open, setOpen] = useState(false)
  const live = useRef({ editor, filePath })
  live.current = { editor, filePath }

  // automatic links run for the open document whether or not the panel is open
  useEffect(
    () =>
      installAutoLinks({
        editor: () => live.current.editor,
        dvh: activeDvhDocs,
        filePath: () => live.current.filePath,
        readSource,
        watch: (paths) => void window.desktop.dvhWatchSources?.(paths),
        writeFields: window.desktop.dvhWriteFields
          ? (request) => window.desktop.dvhWriteFields(request)
          : undefined,
        findSource,
      }),
    [],
  )

  // Read/Write fields of automatic links go back to their workbook shortly after typing
  useEffect(() => {
    if (!editor) return
    const onUpdate = ({ transaction }: { transaction: { docChanged: boolean } }) => {
      if (transaction.docChanged) noteDocumentEdited()
    }
    editor.on('update', onUpdate)
    return () => {
      editor.off('update', onUpdate)
    }
  }, [editor])

  useEffect(() => {
    if (!parsed) {
      setActiveDvhDocs(null)
      return
    }
    let current = true
    void loadDvhDocs(parsed).then((state) => {
      if (!current) return
      setActiveDvhDocs(state)
      refreshAutoWatch()
      void checkLinksOnOpen()
    })
    return () => {
      current = false
    }
  }, [parsed])

  useEffect(() => {
    void window.desktop
      .dvhSystemLocale?.()
      .then(setDvhNumberLocale)
      .catch(() => {})
  }, [])

  useEffect(() => {
    setDvhModelChangeListener(markDirty)
    return () => setDvhModelChangeListener(null)
  }, [markDirty])

  useEffect(() => {
    const toggle = () => setOpen((value) => !value)
    window.addEventListener('dvh:smart-data', toggle)
    // e2e drivers reach the action registry the agent will use (P4 wires it to the agent)
    const w = window as unknown as Record<string, unknown>
    if (w.__genofficeDebugHooks === true) {
      w.__dvhActions = createDocsDvhActions({
        editor: () => live.current.editor,
        filePath: () => live.current.filePath,
        readSource: (path) => window.desktop.dvhReadSource({ path }),
      })
    }
    return () => window.removeEventListener('dvh:smart-data', toggle)
  }, [])

  if (!open || !editor || !parsed) return null
  return (
    <DvhSmartDataPanel
      editor={editor}
      parsed={parsed}
      filePath={filePath}
      onClose={() => setOpen(false)}
    />
  )
}

/** Smart Data of the open document: linked workbooks, fields, insert at the caret, history. */
export function DvhSmartDataPanel({
  editor,
  parsed,
  filePath,
  onClose,
}: {
  editor: Editor
  parsed: ParsedDocFull
  filePath: string | null
  onClose: () => void
}) {
  const { t } = useI18n()
  const [, setTick] = useState(0)
  const repaint = () => setTick((n) => n + 1)
  const [name, setName] = useState('')
  const [value, setValue] = useState('')
  const [error, setError] = useState('')
  const [busy, setBusy] = useState(false)
  /** the history list shows only this field (from a conflict's "History") */
  const [historyField, setHistoryField] = useState<string | null>(null)
  const historyRef = useRef<HTMLDetailsElement>(null)

  const dvh = dvhDocsStateOf(parsed)
  const fields = dvh?.model?.fields ?? []
  const links = dvh?.model?.links ?? []

  // link statuses come from the automatic-link runner; opening the panel re-checks them
  useEffect(() => onLinkStatus(() => setTick((n) => n + 1)), [])
  useEffect(() => {
    void checkLinksOnOpen()
  }, [parsed])

  useEffect(() => {
    // field text edited in the page shows up in the value column
    const onUpdate = () => setTick((n) => n + 1)
    editor.on('update', onUpdate)
    return () => {
      editor.off('update', onUpdate)
    }
  }, [editor])

  const occurrences = fieldOccurrences(editor.state.doc)
  const uses = new Map<string, number>()
  const shownText = new Map<string, string>()
  for (const occ of occurrences) {
    uses.set(occ.fieldId, (uses.get(occ.fieldId) ?? 0) + 1)
    if (!shownText.has(occ.fieldId)) shownText.set(occ.fieldId, occ.text)
  }
  const fieldName = (id: string) => fields.find((f) => f.id === id)?.name ?? id
  const show = (v: unknown) =>
    v === null || ['string', 'number', 'boolean'].includes(typeof v) ? fieldText(v as Scalar) : '…'

  const linkWorkbook = async () => {
    if (!dvh || busy) return
    setBusy(true)
    try {
      const read = await window.desktop.dvhReadSource({ pick: true })
      if (!read) return
      const source = sourceFromRead(read)
      const objects = source ? source.model.fields.length + source.model.collections.length : 0
      if (!source || objects === 0) {
        setError(t('dvhNoSourceModel'))
        return
      }
      const link = linkSource(dvh, source, filePath)
      setLinkStatus(link.id, 'current')
      refreshAutoWatch()
      setError('')
      showToast(t('dvhLinked', { count: objects, name: fileName(read.path) }))
    } catch (e) {
      setError(t('dvhErrReadSource', { error: e instanceof Error ? e.message : String(e) }))
    } finally {
      setBusy(false)
      repaint()
    }
  }

  const update = async (link: DvhLink) => {
    if (!dvh || busy) return
    setBusy(true)
    try {
      const found = await resolveLinkedSource(dvh, link, filePath, readSource, findSource)
      if (found.kind !== 'found') {
        setLinkStatus(link.id, found.kind)
        setError(
          found.kind === 'missing'
            ? t('dvhSourceMissing', { path: link.source.path ?? link.source.relPath ?? '' })
            : '',
        )
        return
      }
      const source = found.source
      if (found.relocated) {
        refreshAutoWatch()
        showToast(t('dvhRelinked', { name: fileName(source.path) }))
      }
      const result = applySource(editor, dvh, link, source, false)
      let tables = result.tables
      let declined = false
      for (const tableId of result.editedTables) {
        const table = dvh.model?.tables.find((tb) => tb.id === tableId)
        if (!table || !window.confirm(t('dvhTableEditedConfirm', { name: table.name }))) {
          declined = true
          continue
        }
        tables += refreshDocTable(editor, dvh, tableId, { force: true, source: 'link' }).refreshed
      }
      if (linkStatus(link.id) !== 'conflict')
        setLinkStatus(link.id, declined ? 'edited' : 'current')
      setError('')
      showToast(t('dvhUpdated', { count: result.fields + tables, name: fileName(source.path) }))
    } finally {
      setBusy(false)
      repaint()
    }
  }

  const reportWrite = (outcome: WriteBackOutcome | undefined, link: DvhLink) => {
    if (!outcome) return
    if (outcome.error) {
      setError(t('dvhWriteFailed', { error: outcome.error }))
      return
    }
    setError(
      outcome.refused.length > 0 ? t('dvhWriteRefused', { count: outcome.refused.length }) : '',
    )
    if (outcome.written > 0) {
      const name = fileName(link.source.path ?? link.source.relPath ?? link.source.docId)
      showToast(t('dvhWritten', { count: outcome.written, name }))
    }
  }

  const writeBack = async (link: DvhLink) => {
    if (busy) return
    setBusy(true)
    try {
      reportWrite(await writeBackLink(link.id), link)
    } finally {
      setBusy(false)
      repaint()
    }
  }

  const settle = async (link: DvhLink, fieldId: string, keep: 'mine' | 'source') => {
    if (busy) return
    setBusy(true)
    try {
      reportWrite(await resolveConflict(link.id, fieldId, keep), link)
    } finally {
      setBusy(false)
      repaint()
    }
  }

  const showHistory = (fieldId: string) => {
    setHistoryField(fieldId)
    if (historyRef.current) historyRef.current.open = true
    historyRef.current?.scrollIntoView({ block: 'nearest' })
  }

  /** Points the link at another copy of its workbook: same docId, or one carrying every linked object. */
  const changeSource = async (link: DvhLink, picked?: string) => {
    if (!dvh || busy) return
    setBusy(true)
    try {
      const read = picked
        ? await window.desktop.dvhReadSource({ path: picked })
        : await window.desktop.dvhReadSource({ pick: true })
      if (!read) return
      const source = sourceFromRead(read)
      const ids = new Set([
        ...(source?.model.fields.map((f) => f.id) ?? []),
        ...(source?.model.collections.map((c) => c.id) ?? []),
      ])
      if (
        !source ||
        (source.model.docId !== link.source.docId && !link.targets.every((id) => ids.has(id)))
      ) {
        setError(t('dvhNotThisSource'))
        return
      }
      relocateLink(dvh, link, source.path, filePath, source.model.docId)
      refreshAutoWatch()
      applySource(editor, dvh, link, source, false)
      setError('')
      showToast(t('dvhRelinked', { name: fileName(source.path) }))
    } catch (e) {
      setError(t('dvhErrReadSource', { error: e instanceof Error ? e.message : String(e) }))
    } finally {
      setBusy(false)
      repaint()
    }
  }

  const openSource = async (link: DvhLink) => {
    const found = await resolveLinkedSource(dvh!, link, filePath, readSource, findSource)
    const path = found.kind === 'found' ? found.source.path : null
    if (!path || !(await window.desktop.dvhOpenSource?.(path))) {
      setError(
        t('dvhOpenSourceFailed', { path: path ?? link.source.path ?? link.source.relPath ?? '' }),
      )
    }
    repaint()
  }

  const changeMode = (link: DvhLink, mode: LinkUpdateMode) => {
    if (!dvh) return
    setLinkUpdateMode(dvh, link.id, mode)
    refreshAutoWatch()
    // switching to on-open/automatic catches up right away
    if (mode !== 'manual') void checkLinksOnOpen()
    repaint()
  }

  const unlink = (link: DvhLink) => {
    if (!dvh) return
    const label = fileName(link.source.path ?? link.source.relPath ?? link.source.docId)
    if (!window.confirm(t('dvhUnlinkConfirm', { name: label }))) return
    unlinkSource(editor, dvh, link.id)
    refreshAutoWatch()
    repaint()
  }

  const insert = (fieldId: string) => {
    const field = dvh?.model?.fields.find((f) => f.id === fieldId)
    if (!dvh || !field) return
    insertSmartField(editor, dvh, field)
    repaint()
  }

  const addField = () => {
    if (!dvh) return
    const trimmed = name.trim()
    if (!trimmed || fields.some((f) => f.name === trimmed)) {
      setError(t('dvhErrName'))
      return
    }
    insertSmartField(editor, dvh, {
      id: newDvhId('f'),
      name: trimmed,
      type: 'text',
      value: value === '' ? null : value,
      access: 'readwrite',
    })
    setName('')
    setValue('')
    setError('')
    repaint()
  }

  const history = historyEntries(dvh)
    .filter((cs) => !historyField || cs.changes.some((c) => c.objectId === historyField))
    .slice(0, HISTORY_ROWS)
  const canWrite = typeof window.desktop.dvhWriteFields === 'function'

  return (
    <div className="dvh-docs-panel" role="dialog" aria-label={t('dvhPanelTitle')}>
      <div className="dvh-docs-panel-head">
        <strong>{t('dvhPanelTitle')}</strong>
        <button type="button" onClick={onClose} aria-label={t('dvhClose')}>
          ×
        </button>
      </div>

      <div className="dvh-docs-panel-section">
        <div className="dvh-docs-panel-subhead">
          <span>{t('dvhLinks')}</span>
          <button type="button" disabled={busy} onClick={() => void linkWorkbook()}>
            {t('dvhLinkWorkbook')}
          </button>
        </div>
        {links.length === 0 ? (
          <p className="dvh-docs-panel-empty">{t('dvhNoLinks')}</p>
        ) : (
          links.map((link) => (
            <div key={link.id} className="dvh-docs-panel-link" data-link-id={link.id}>
              <div className="dvh-docs-panel-link-main">
                <span className="dvh-docs-panel-link-name" title={link.source.path}>
                  {fileName(link.source.path ?? link.source.relPath ?? link.source.docId)}
                </span>
                {link.lastSync ? (
                  <span className="dvh-docs-panel-muted">
                    {t('dvhLastSync', { time: new Date(link.lastSync.at).toLocaleString() })}
                  </span>
                ) : null}
                {linkStatus(link.id) === 'stale' ? (
                  <span className="dvh-docs-panel-warn" data-status="stale">
                    {t('dvhStale')}
                  </span>
                ) : linkStatus(link.id) === 'missing' ? (
                  <span className="dvh-docs-panel-warn" data-status="missing">
                    {t('dvhSourceMissing', {
                      path: link.source.path ?? link.source.relPath ?? '',
                    })}
                  </span>
                ) : linkStatus(link.id) === 'edited' ? (
                  <span className="dvh-docs-panel-warn" data-status="edited">
                    {t('dvhStatusEdited')}
                  </span>
                ) : linkStatus(link.id) === 'conflict' ? (
                  <span className="dvh-docs-panel-warn" data-status="conflict">
                    {t('dvhStatusConflict')}
                  </span>
                ) : linkStatus(link.id) === 'ambiguous' ? (
                  <span className="dvh-docs-panel-warn" data-status="ambiguous">
                    {t('dvhAmbiguous')}
                  </span>
                ) : null}
                {ambiguousSources(link.id).length > 0 ? (
                  <ul className="dvh-docs-panel-choices">
                    {ambiguousSources(link.id).map((path) => (
                      <li key={path}>
                        <span title={path}>{path}</span>
                        <button
                          type="button"
                          disabled={busy}
                          onClick={() => void changeSource(link, path)}
                        >
                          {t('dvhUseThis')}
                        </button>
                      </li>
                    ))}
                  </ul>
                ) : null}
                {linkConflicts(link.id).map((c) => (
                  <div
                    key={c.fieldId}
                    className="dvh-docs-panel-conflict"
                    data-field-id={c.fieldId}
                  >
                    <strong>{fieldName(c.fieldId)}</strong>
                    <div>
                      {t('dvhConflictHere')}: <ins>{c.local}</ins>
                    </div>
                    <div>
                      {t('dvhConflictSource')}: <ins>{c.source.text}</ins>
                    </div>
                    <div className="dvh-docs-panel-link-actions">
                      <button
                        type="button"
                        disabled={busy || !canWrite}
                        onClick={() => void settle(link, c.fieldId, 'mine')}
                      >
                        {t('dvhKeepMine')}
                      </button>
                      <button
                        type="button"
                        disabled={busy}
                        onClick={() => void settle(link, c.fieldId, 'source')}
                      >
                        {t('dvhKeepSource')}
                      </button>
                      <button type="button" onClick={() => showHistory(c.fieldId)}>
                        {t('dvhViewHistory')}
                      </button>
                    </div>
                  </div>
                ))}
                {linkPendingWrites(editor, dvh!, link).length > 0 ? (
                  <span className="dvh-docs-panel-muted" data-status="pending-writes">
                    {t('dvhPendingWrites', { count: linkPendingWrites(editor, dvh!, link).length })}
                  </span>
                ) : null}
                <div className="dvh-docs-panel-link-actions">
                  <select
                    value={link.update}
                    aria-label={t('dvhUpdateMode')}
                    onChange={(event) => changeMode(link, event.target.value as LinkUpdateMode)}
                  >
                    <option value="manual">{t('dvhModeManual')}</option>
                    <option value="onOpen">{t('dvhModeOnOpen')}</option>
                    <option value="auto">{t('dvhModeAuto')}</option>
                  </select>
                  <button type="button" disabled={busy} onClick={() => void update(link)}>
                    {t('dvhUpdateFromSource')}
                  </button>
                  <button
                    type="button"
                    data-action="write-back"
                    disabled={
                      busy || !canWrite || linkPendingWrites(editor, dvh!, link).length === 0
                    }
                    onClick={() => void writeBack(link)}
                  >
                    {t('dvhWriteBack')}
                  </button>
                  <button type="button" disabled={busy} onClick={() => void openSource(link)}>
                    {t('dvhOpenSource')}
                  </button>
                  <button type="button" disabled={busy} onClick={() => void changeSource(link)}>
                    {t('dvhChangeSource')}
                  </button>
                  <button type="button" onClick={() => unlink(link)}>
                    {t('dvhUnlink')}
                  </button>
                </div>
              </div>
            </div>
          ))
        )}
      </div>

      <div className="dvh-docs-panel-section">
        {fields.length === 0 ? (
          <p className="dvh-docs-panel-empty">{t('dvhEmpty')}</p>
        ) : (
          <table className="dvh-docs-panel-table">
            <thead>
              <tr>
                <th>{t('dvhColName')}</th>
                <th>{t('dvhColValue')}</th>
                <th>{t('dvhColUses')}</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {fields.map((field) => (
                <tr key={field.id} data-field-id={field.id}>
                  <td>{field.name}</td>
                  <td>{shownText.get(field.id) ?? fieldText(field.value)}</td>
                  <td>{uses.get(field.id) ?? 0}</td>
                  <td>
                    <button type="button" onClick={() => insert(field.id)}>
                      {t('dvhInsert')}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        <div className="dvh-docs-panel-new">
          <input
            value={name}
            placeholder={t('dvhNamePlaceholder')}
            aria-label={t('dvhNamePlaceholder')}
            onChange={(event) => setName(event.target.value)}
          />
          <input
            value={value}
            placeholder={t('dvhValuePlaceholder')}
            aria-label={t('dvhValuePlaceholder')}
            onChange={(event) => setValue(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') addField()
            }}
          />
          <button type="button" onClick={addField}>
            {t('dvhAddField')}
          </button>
        </div>
        {error ? <p className="dvh-docs-panel-error">{error}</p> : null}
      </div>

      <DvhDocsTablesSection editor={editor} dvh={dvh} onChange={repaint} />

      <details ref={historyRef} className="dvh-docs-panel-section dvh-docs-panel-history">
        <summary>
          {historyField ? t('dvhHistoryFor', { name: fieldName(historyField) }) : t('dvhHistory')}
        </summary>
        {historyField ? (
          <button type="button" onClick={() => setHistoryField(null)}>
            {t('dvhHistoryAll')}
          </button>
        ) : null}
        {history.length === 0 ? (
          <p className="dvh-docs-panel-empty">{t('dvhNoHistory')}</p>
        ) : (
          <ul>
            {history.map((cs) => (
              <li key={cs.id} data-action={cs.action}>
                <span className="dvh-docs-panel-muted">{new Date(cs.at).toLocaleString()}</span>{' '}
                <code>{cs.action}</code>
                {cs.changes
                  .filter((c) => c.path === 'value')
                  .map((c) => (
                    <div key={`${cs.id}:${c.objectId}`} className="dvh-docs-panel-diff">
                      {fieldName(c.objectId)}: <del>{show(c.before)}</del> →{' '}
                      <ins>{show(c.after)}</ins>
                    </div>
                  ))}
              </li>
            ))}
          </ul>
        )}
      </details>
    </div>
  )
}

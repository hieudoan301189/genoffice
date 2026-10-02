import { useCallback, useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { ParsedDocFull } from '@genoffice/docx-engine'
import { fieldText, newDvhId, type DvhLink, type Scalar } from '@genoffice/dvh-model'
import { useI18n } from '../i18n/locale'
import { showToast } from './toast-bus'
import { createDocsDvhActions } from '../dvh-actions'
import {
  dvhDocsStateOf,
  fieldOccurrences,
  historyEntries,
  insertSmartField,
  linkIsStale,
  linkSource,
  linkSourcePaths,
  loadDvhDocs,
  setActiveDvhDocs,
  setDvhModelChangeListener,
  sourceFromRead,
  type DvhSource,
} from '../dvh-smart-data'
import { refreshDocTable, setDvhNumberLocale, updateLinkFromSource } from '../dvh-tables'
import { DvhDocsTablesSection } from './DvhDocsTablesSection'
import './dvh-smart-data.css'

const fileName = (path: string) => path.split(/[\\/]/).pop() ?? path
const HISTORY_ROWS = 20

/** Reads the first candidate path that still holds the linked workbook. */
async function readLinkedSource(link: DvhLink, docPath: string | null): Promise<DvhSource | null> {
  for (const path of linkSourcePaths(link, docPath)) {
    try {
      const source = sourceFromRead(await window.desktop.dvhReadSource({ path }))
      if (source && source.model.docId === link.source.docId) return source
    } catch {
      // moved or unreadable: try the next candidate
    }
  }
  return null
}

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

  useEffect(() => {
    if (!parsed) {
      setActiveDvhDocs(null)
      return
    }
    let current = true
    void loadDvhDocs(parsed).then((state) => {
      if (current) setActiveDvhDocs(state)
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

type LinkStatus = 'current' | 'stale' | 'missing'

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
  const [status, setStatus] = useState<Record<string, LinkStatus>>({})

  const dvh = dvhDocsStateOf(parsed)
  const fields = dvh?.model?.fields ?? []
  const links = dvh?.model?.links ?? []

  /** compares each link's last pull with the workbook on disk (stale-value warning) */
  const checkLinks = useCallback(async () => {
    const state = await loadDvhDocs(parsed)
    const next: Record<string, LinkStatus> = {}
    for (const link of state.model?.links ?? []) {
      const source = await readLinkedSource(link, filePath)
      next[link.id] = !source ? 'missing' : linkIsStale(link, source) ? 'stale' : 'current'
    }
    setStatus(next)
  }, [parsed, filePath])

  useEffect(() => {
    void checkLinks()
  }, [checkLinks])

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
      linkSource(dvh, source, filePath)
      setError('')
      showToast(t('dvhLinked', { count: objects, name: fileName(read.path) }))
      await checkLinks()
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
      const source = await readLinkedSource(link, filePath)
      if (!source) {
        setStatus((s) => ({ ...s, [link.id]: 'missing' }))
        setError(t('dvhSourceMissing', { path: link.source.path ?? link.source.relPath ?? '' }))
        return
      }
      const result = updateLinkFromSource(editor, dvh, link, source)
      let tables = result.tables
      for (const tableId of result.editedTables) {
        const table = dvh.model?.tables.find((tb) => tb.id === tableId)
        if (!table || !window.confirm(t('dvhTableEditedConfirm', { name: table.name }))) continue
        tables += refreshDocTable(editor, dvh, tableId, { force: true, source: 'link' }).refreshed
      }
      setStatus((s) => ({ ...s, [link.id]: 'current' }))
      setError('')
      showToast(t('dvhUpdated', { count: result.fields + tables, name: fileName(source.path) }))
    } finally {
      setBusy(false)
      repaint()
    }
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

  const history = historyEntries(dvh).slice(0, HISTORY_ROWS)

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
                {status[link.id] === 'stale' ? (
                  <span className="dvh-docs-panel-warn" data-status="stale">
                    {t('dvhStale')}
                  </span>
                ) : status[link.id] === 'missing' ? (
                  <span className="dvh-docs-panel-warn" data-status="missing">
                    {t('dvhSourceMissing', {
                      path: link.source.path ?? link.source.relPath ?? '',
                    })}
                  </span>
                ) : null}
              </div>
              <button type="button" disabled={busy} onClick={() => void update(link)}>
                {t('dvhUpdateFromSource')}
              </button>
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

      <details className="dvh-docs-panel-section dvh-docs-panel-history">
        <summary>{t('dvhHistory')}</summary>
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

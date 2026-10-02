import { useEffect, useState } from 'react'
import type { DvhProjectInfo } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { showToast } from './toast-bus'
import { linkSource, sourceFromRead, type DvhDocsState } from '../dvh-smart-data'
import { refreshAutoWatch, setLinkStatus } from '../dvh-auto'
import { noteUnrecordable } from '../dvh-workflow'

type Row = { id: string; values: Record<string, string | number | boolean | null> }

const text = (v: string | number | boolean | null | undefined) =>
  v === null || v === undefined ? '' : String(v)

/**
 * Project data of the Smart Data panel (P10): the schema packs of the
 * document's project, its objects by type (edited in place), what the pack
 * rules find, a link from this document to the project's objects, and the
 * two-way exchange with QLCL-DVH with its conflicts.
 */
export function DvhProjectSection({
  dvh,
  filePath,
  onChange,
}: {
  dvh: DvhDocsState | null
  filePath: string | null
  onChange: () => void
}) {
  const { t } = useI18n()
  const [info, setInfo] = useState<DvhProjectInfo | null>(null)
  const [type, setType] = useState('')
  const [rows, setRows] = useState<Row[]>([])
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const api = window.desktop

  const reload = async (nextType = type) => {
    if (!api.dvhProjectInfo) return
    const next = await api.dvhProjectInfo(filePath)
    setInfo(next)
    const chosen = next.types.find((x) => x.name === nextType) ?? next.types[0]
    setType(chosen?.name ?? '')
    setRows(chosen ? await api.dvhProjectObjects(filePath, chosen.name) : [])
  }

  useEffect(() => {
    void reload().catch((e: unknown) => setMessage(e instanceof Error ? e.message : String(e)))
    // reload only when the document (its project) changes
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [filePath])

  const run = async (job: () => Promise<void>) => {
    if (busy) return
    setBusy(true)
    try {
      await job()
      setMessage('')
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
      onChange()
    }
  }

  if (!api.dvhProjectInfo) return null
  const current = info?.types.find((x) => x.name === type)

  const installPack = () =>
    run(async () => {
      const next = await api.dvhProjectInstallPack(filePath)
      if (next) {
        setInfo(next)
        await reload()
      }
    })

  const setValue = (row: Row | null, key: string, value: string) =>
    run(async () => {
      if (!current) return
      if (row && text(row.values[key]) === value) return
      await api.dvhProjectSet(filePath, {
        type: current.name,
        objectId: row?.id ?? null,
        values: { [key]: value === '' ? null : value },
      })
      await reload(current.name)
    })

  const remove = (row: Row) =>
    run(async () => {
      if (
        !window.confirm(
          t('dvhProjectDeleteConfirm', {
            name: text(row.values[current?.fields[0]?.key ?? '']) || row.id,
          }),
        )
      )
        return
      await api.dvhProjectDelete(filePath, row.id)
      await reload(type)
    })

  const link = () =>
    run(async () => {
      if (!dvh || !info) return
      noteUnrecordable(t('dvhProjectLink'))
      const source = sourceFromRead(await api.dvhReadSource({ path: info.uri }))
      if (!source) {
        setMessage(t('dvhProjectNoData'))
        return
      }
      const made = linkSource(dvh, source, filePath)
      setLinkStatus(made.id, 'current')
      refreshAutoWatch()
      showToast(
        t('dvhProjectLinked', {
          count: source.model.fields.length + source.model.collections.length,
        }),
      )
    })

  const sync = (partner: string | null) =>
    run(async () => {
      noteUnrecordable(t('dvhProjectSync'))
      const result = await api.dvhProjectSync(filePath, partner)
      if (!result) return
      if ('error' in result) {
        setMessage(result.error)
        return
      }
      showToast(
        t('dvhProjectSynced', {
          pulled: result.stats.pulled ?? 0,
          pushed: result.stats.pushed ?? 0,
          created: result.stats.created ?? 0,
          deleted: result.stats.deleted ?? 0,
          conflicts: result.conflicts.length,
        }),
      )
      await reload()
    })

  const settle = (partner: string, index: number, keep: 'mine' | 'theirs') =>
    run(async () => {
      await api.dvhProjectResolve(filePath, partner, index, keep)
      await reload()
    })

  return (
    <details className="dvh-docs-panel-section" data-section="project">
      <summary>
        {t('dvhProject')}
        {info ? ` — ${info.name}` : ''}
      </summary>
      <div className="dvh-docs-panel-subhead">
        <span>{t('dvhProjectPacks')}</span>
        <button type="button" disabled={busy} onClick={() => void installPack()}>
          {t('dvhProjectInstallPack')}
        </button>
      </div>
      {info && info.packs.length === 0 ? (
        <p className="dvh-docs-panel-empty">{t('dvhProjectNoPack')}</p>
      ) : (
        <ul className="dvh-docs-panel-choices">
          {info?.packs.map((p) => (
            <li key={p.id}>
              {p.name} <span className="dvh-docs-panel-muted">{p.packVersion}</span>
            </li>
          ))}
        </ul>
      )}

      {info && info.types.length > 0 ? (
        <>
          <div className="dvh-docs-panel-new">
            <select
              aria-label={t('dvhProjectType')}
              value={type}
              onChange={(event) => void run(() => reload(event.target.value))}
            >
              {info.types.map((x) => (
                <option key={x.name} value={x.name}>
                  {x.label ?? x.name} ({x.count})
                </option>
              ))}
            </select>
            <button
              type="button"
              disabled={busy || !current || (current.single && rows.length > 0)}
              onClick={() => current && void setValue(null, current.fields[0]!.key, '')}
            >
              {t('dvhProjectAdd')}
            </button>
          </div>
          {current ? (
            <div className="dvh-docs-panel-scroll">
              <table className="dvh-docs-panel-table" aria-label={current.label ?? current.name}>
                <thead>
                  <tr>
                    {current.fields.map((f) => (
                      <th key={f.key}>{f.label ?? f.key}</th>
                    ))}
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row) => (
                    <tr key={row.id} data-object-id={row.id}>
                      {current.fields.map((f) => (
                        <td key={f.key}>
                          {f.enumValues ? (
                            <select
                              aria-label={f.label ?? f.key}
                              value={text(row.values[f.key])}
                              onChange={(event) => void setValue(row, f.key, event.target.value)}
                            >
                              <option value="" />
                              {f.enumValues.map((v) => (
                                <option key={v} value={v}>
                                  {v}
                                </option>
                              ))}
                            </select>
                          ) : (
                            <input
                              key={text(row.values[f.key])}
                              className="dvh-docs-panel-value"
                              aria-label={f.label ?? f.key}
                              defaultValue={text(row.values[f.key])}
                              onBlur={(event) => void setValue(row, f.key, event.target.value)}
                              onKeyDown={(event) => {
                                if (event.key === 'Enter') event.currentTarget.blur()
                              }}
                            />
                          )}
                        </td>
                      ))}
                      <td>
                        <button
                          type="button"
                          aria-label={t('dvhProjectDelete')}
                          onClick={() => void remove(row)}
                        >
                          ×
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : null}
        </>
      ) : null}

      {info && info.issues.length > 0 ? (
        <details className="dvh-docs-panel-history">
          <summary>{t('dvhProjectIssues', { count: info.issues.length })}</summary>
          <ul className="dvh-docs-panel-choices">
            {info.issues.slice(0, 20).map((issue, i) => (
              <li
                key={`${issue.objectId}-${issue.rule ?? issue.field ?? i}`}
                className={
                  issue.severity === 'error' ? 'dvh-docs-panel-warn' : 'dvh-docs-panel-muted'
                }
              >
                {issue.message}
              </li>
            ))}
          </ul>
        </details>
      ) : null}

      <div className="dvh-docs-panel-link-actions">
        <button type="button" disabled={busy || !dvh || !info} onClick={() => void link()}>
          {t('dvhProjectLink')}
        </button>
        <button
          type="button"
          disabled={busy || !info?.packs.length}
          onClick={() => void sync(null)}
        >
          {t('dvhProjectSync')}
        </button>
      </div>
      {info?.partners.map((partner) => (
        <div key={partner.path} className="dvh-docs-panel-link" data-partner={partner.path}>
          <div className="dvh-docs-panel-link-main">
            <span className="dvh-docs-panel-link-name" title={partner.path}>
              {partner.path.split(/[\\/]/).pop()}
            </span>
            <span className="dvh-docs-panel-muted">
              {t('dvhProjectLastSync', { time: new Date(partner.at).toLocaleString() })}
            </span>
            {partner.conflicts.map((c, i) => (
              <div key={`${c.objectId}-${c.key ?? c.kind}`} className="dvh-docs-panel-conflict">
                <strong>
                  {c.type} {c.objectId}
                  {c.key ? `.${c.key}` : ''}
                </strong>
                {c.kind === 'value' ? (
                  <>
                    <div>
                      {t('dvhConflictHere')}: <ins>{c.local}</ins>
                    </div>
                    <div>
                      {t('dvhConflictSource')}: <ins>{c.remote}</ins>
                    </div>
                  </>
                ) : (
                  <div>
                    {t(
                      c.kind === 'deleted-there'
                        ? 'dvhProjectDeletedThere'
                        : 'dvhProjectDeletedHere',
                    )}
                  </div>
                )}
                <div className="dvh-docs-panel-link-actions">
                  <button
                    type="button"
                    disabled={busy}
                    onClick={() => void settle(partner.path, i, 'mine')}
                  >
                    {t('dvhKeepMine')}
                  </button>
                  <button
                    type="button"
                    disabled={busy}
                    onClick={() => void settle(partner.path, i, 'theirs')}
                  >
                    {t('dvhKeepSource')}
                  </button>
                </div>
              </div>
            ))}
            <div className="dvh-docs-panel-link-actions">
              <button type="button" disabled={busy} onClick={() => void sync(partner.path)}>
                {t('dvhProjectSyncAgain')}
              </button>
            </div>
          </div>
        </div>
      ))}
      {message ? <p className="dvh-docs-panel-error">{message}</p> : null}
    </details>
  )
}

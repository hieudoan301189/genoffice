import { useState } from 'react'
import { useI18n } from './i18n/locale'
import { dvhSheetsContext, dvhStateOf } from './dvh-smart-data'
import {
  collectionRanges,
  createCollection,
  createTable,
  renderTableInSheet,
  setTableStyle,
  tableRanges,
} from './dvh-tables'

const a1 = (row: number, column: number) => {
  let n = column + 1
  let letters = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    letters = String.fromCharCode(65 + rem) + letters
    n = Math.floor((n - 1) / 26)
  }
  return `${letters}${row + 1}`
}

/** Collections (ranges with a title row) and the DVH tables rendered from them. */
export function DvhTablesSection({ onChange }: { onChange: () => void }) {
  const { t } = useI18n()
  const [name, setName] = useState('')
  const [error, setError] = useState('')
  const ctx = dvhSheetsContext()
  const runtime = ctx?.getRuntime() ?? null
  const dvh = dvhStateOf(ctx?.getState())
  const collections = dvh ? collectionRanges(runtime, dvh) : []
  const tables = dvh ? tableRanges(runtime, dvh) : []

  const selection = () => {
    const workbook = runtime?.univerAPI.getActiveWorkbook()
    const range = workbook?.getActiveRange()
    if (!workbook || !range) return null
    return {
      sheetName: workbook.getActiveSheet().getSheetName(),
      row: range.getRow(),
      column: range.getColumn(),
      rows: range.getHeight(),
      columns: range.getWidth(),
    }
  }

  const done = () => {
    ctx?.markPending()
    setError('')
    onChange()
    // typing after an action goes to the grid, not into this panel
    ;(document.activeElement as HTMLElement | null)?.blur()
  }

  const makeCollection = () => {
    if (!ctx || !runtime || !dvh) return
    const rect = selection()
    if (!rect || rect.rows < 2) {
      setError(t('dvhErrSelectRange'))
      return
    }
    try {
      createCollection(runtime, dvh, { name, rect })
      setName('')
      done()
    } catch {
      setError(t('dvhErrCollectionName'))
    }
  }

  const insertTable = (collectionId: string, collectionName: string) => {
    if (!ctx || !runtime || !dvh) return
    const rect = selection()
    if (!rect) {
      setError(t('dvhErrSelectAnchor'))
      return
    }
    const taken = new Set(dvh.model?.tables.map((tb) => tb.name) ?? [])
    let tableName = collectionName
    for (let n = 2; taken.has(tableName); n++) tableName = `${collectionName} ${n}`
    const table = createTable(dvh, { name: tableName, collectionId })
    const result = renderTableInSheet(runtime, dvh, table.id, { anchor: rect })
    ctx.notify(
      t('dvhTableRendered', {
        name: table.name,
        cell: a1(rect.row, rect.column),
        rows: result.rows,
      }),
    )
    done()
  }

  const refresh = (tableId: string, tableName: string) => {
    if (!ctx || !runtime || !dvh) return
    if (renderTableInSheet(runtime, dvh, tableId).needsConfirm) {
      if (!window.confirm(t('dvhTableEditedConfirm', { name: tableName }))) return
      renderTableInSheet(runtime, dvh, tableId, { force: true })
    }
    done()
  }

  const changeMode = (tableId: string, tableName: string, mode: 'source' | 'destination') => {
    if (!dvh) return
    setTableStyle(dvh, tableId, { mode })
    refresh(tableId, tableName)
  }

  return (
    <>
      <div className="dvh-smart-data-subhead">
        <strong>{t('dvhCollections')}</strong>
      </div>
      {collections.length === 0 ? (
        <p className="dvh-smart-data-empty">{t('dvhNoCollections')}</p>
      ) : (
        <table className="dvh-smart-data-table" aria-label={t('dvhCollections')}>
          <thead>
            <tr>
              <th>{t('dvhColName')}</th>
              <th>{t('dvhColRange')}</th>
              <th>{t('dvhColRows')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {collections.map(({ collection, formula }) => (
              <tr key={collection.id} data-collection-id={collection.id}>
                <td>{collection.name}</td>
                <td className={formula ? undefined : 'dvh-smart-data-broken'}>
                  {formula ?? t('dvhBroken')}
                </td>
                <td>{collection.rows.length}</td>
                <td>
                  <button type="button" onClick={() => insertTable(collection.id, collection.name)}>
                    {t('dvhInsertTable')}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      <div className="dvh-smart-data-new">
        <input
          value={name}
          placeholder={t('dvhCollectionNamePlaceholder')}
          aria-label={t('dvhCollectionNamePlaceholder')}
          onChange={(event) => setName(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') makeCollection()
          }}
        />
        <button type="button" onClick={makeCollection}>
          {t('dvhCollectionFromSelection')}
        </button>
      </div>

      <div className="dvh-smart-data-subhead">
        <strong>{t('dvhTables')}</strong>
      </div>
      {tables.length === 0 ? (
        <p className="dvh-smart-data-empty">{t('dvhNoTables')}</p>
      ) : (
        <table className="dvh-smart-data-table" aria-label={t('dvhTables')}>
          <tbody>
            {tables.map(({ table, formula }) => (
              <tr key={table.id} data-table-id={table.id}>
                <td>{table.name}</td>
                <td>{formula ?? '—'}</td>
                <td>
                  <select
                    value={table.style.mode}
                    aria-label={table.name}
                    onChange={(event) =>
                      changeMode(
                        table.id,
                        table.name,
                        event.target.value as 'source' | 'destination',
                      )
                    }
                  >
                    <option value="source">{t('dvhStyleSource')}</option>
                    <option value="destination">{t('dvhStyleDestination')}</option>
                  </select>
                </td>
                <td>
                  <button
                    type="button"
                    disabled={!formula}
                    onClick={() => refresh(table.id, table.name)}
                  >
                    {t('dvhRefresh')}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {error ? <p className="dvh-smart-data-error">{error}</p> : null}
    </>
  )
}

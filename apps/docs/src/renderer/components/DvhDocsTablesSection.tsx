import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import type { DvhDocsState } from '../dvh-smart-data'
import {
  createDocTable,
  insertDocTable,
  refreshDocTable,
  setDocTableStyle,
  tableOccurrences,
} from '../dvh-tables'

/** Linked collections (insert as a DVH.Table) and the tables already in the document. */
export function DvhDocsTablesSection({
  editor,
  dvh,
  onChange,
}: {
  editor: Editor
  dvh: DvhDocsState | null
  onChange: () => void
}) {
  const { t } = useI18n()
  const collections = dvh?.model?.collections ?? []
  const tables = dvh?.model?.tables ?? []
  const shown = new Set(tableOccurrences(editor.state.doc).map((o) => o.tableId))

  const insert = (collectionId: string) => {
    if (!dvh) return
    const table = createDocTable(dvh, { collectionId })
    insertDocTable(editor, dvh, table)
    onChange()
  }

  const refresh = (tableId: string, tableName: string) => {
    if (!dvh) return
    if (refreshDocTable(editor, dvh, tableId).needsConfirm) {
      if (!window.confirm(t('dvhTableEditedConfirm', { name: tableName }))) return
      refreshDocTable(editor, dvh, tableId, { force: true })
    }
    onChange()
  }

  const changeMode = (tableId: string, tableName: string, mode: 'source' | 'destination') => {
    if (!dvh) return
    setDocTableStyle(dvh, tableId, { mode })
    refresh(tableId, tableName)
  }

  return (
    <div className="dvh-docs-panel-section">
      <div className="dvh-docs-panel-subhead">
        <span>{t('dvhCollections')}</span>
      </div>
      {collections.length === 0 ? (
        <p className="dvh-docs-panel-empty">{t('dvhNoCollections')}</p>
      ) : (
        <table className="dvh-docs-panel-table" aria-label={t('dvhCollections')}>
          <tbody>
            {collections.map((collection) => (
              <tr key={collection.id} data-collection-id={collection.id}>
                <td>{collection.name}</td>
                <td>{t('dvhRowCount', { count: collection.rows.length })}</td>
                <td>
                  <button type="button" onClick={() => insert(collection.id)}>
                    {t('dvhInsertTable')}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      {tables.length > 0 ? (
        <>
          <div className="dvh-docs-panel-subhead">
            <span>{t('dvhTables')}</span>
          </div>
          <table className="dvh-docs-panel-table" aria-label={t('dvhTables')}>
            <tbody>
              {tables.map((table) => (
                <tr key={table.id} data-table-id={table.id}>
                  <td>{table.name}</td>
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
                      disabled={!shown.has(table.id)}
                      onClick={() => refresh(table.id, table.name)}
                    >
                      {t('dvhRefresh')}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      ) : null}
    </div>
  )
}

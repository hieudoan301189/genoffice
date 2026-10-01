import { useEffect, useState } from 'react'
import { fieldText } from '@genoffice/dvh-model'
import { useI18n } from './i18n/locale'
import {
  boundCellValue,
  createField,
  dvhBindings,
  dvhSheetsContext,
  dvhStateOf,
  removeField,
} from './dvh-smart-data'

/** Smart Data fields of the open workbook: list, bind the selected cell, remove. */
export function DvhSmartDataPanel({ onClose }: { onClose: () => void }) {
  const { t } = useI18n()
  const [, setTick] = useState(0)
  const [name, setName] = useState('')
  const [error, setError] = useState('')
  // bound cells change under the panel (typing, cut/paste): repaint while open
  useEffect(() => {
    const id = setInterval(() => setTick((n) => n + 1), 1000)
    return () => clearInterval(id)
  }, [])

  const ctx = dvhSheetsContext()
  const runtime = ctx?.getRuntime() ?? null
  const dvh = dvhStateOf(ctx?.getState())
  const bindings = runtime && dvh ? dvhBindings(runtime, dvh) : []

  const bind = () => {
    if (!ctx || !runtime || !dvh) {
      setError(t('dvhErrNoFile'))
      return
    }
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const range = workbook?.getActiveRange()
    if (!workbook || !range || range.getWidth() !== 1 || range.getHeight() !== 1) {
      setError(t('dvhErrSelectCell'))
      return
    }
    try {
      const field = createField(runtime, dvh, {
        name,
        sheetName: workbook.getActiveSheet().getSheetName(),
        row: range.getRow(),
        column: range.getColumn(),
      })
      ctx.markPending()
      ctx.notify(t('dvhBound', { name: field.name, cell: range.getA1Notation() }))
      setName('')
      setError('')
      // typing after binding goes to the grid, not into this input
      ;(document.activeElement as HTMLElement | null)?.blur()
    } catch {
      setError(t('dvhErrName'))
    }
  }

  const remove = (fieldId: string) => {
    if (!ctx || !runtime || !dvh) return
    removeField(runtime, dvh, fieldId)
    ctx.markPending()
    setTick((n) => n + 1)
  }

  return (
    <div className="dvh-smart-data" role="dialog" aria-label={t('dvhPanelTitle')}>
      <div className="dvh-smart-data-head">
        <strong>{t('dvhPanelTitle')}</strong>
        <button type="button" onClick={onClose} aria-label={t('dvhClose')}>
          ×
        </button>
      </div>
      {bindings.length === 0 ? (
        <p className="dvh-smart-data-empty">{t('dvhEmpty')}</p>
      ) : (
        <table className="dvh-smart-data-table">
          <thead>
            <tr>
              <th>{t('dvhColName')}</th>
              <th>{t('dvhColCell')}</th>
              <th>{t('dvhColValue')}</th>
              <th />
            </tr>
          </thead>
          <tbody>
            {bindings.map(({ field, formula, broken }) => (
              <tr key={field.id} data-field-id={field.id}>
                <td>{field.name}</td>
                <td className={broken ? 'dvh-smart-data-broken' : undefined}>
                  {broken ? t('dvhBroken') : formula}
                </td>
                <td>
                  {fieldText(
                    !broken && formula && runtime
                      ? (boundCellValue(runtime, formula) ?? field.value)
                      : field.value,
                  )}
                </td>
                <td>
                  <button type="button" onClick={() => remove(field.id)}>
                    {t('dvhRemove')}
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
          placeholder={t('dvhNamePlaceholder')}
          aria-label={t('dvhNamePlaceholder')}
          onChange={(event) => setName(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') bind()
          }}
        />
        <button type="button" onClick={bind}>
          {t('dvhBindCell')}
        </button>
      </div>
      {error ? <p className="dvh-smart-data-error">{error}</p> : null}
    </div>
  )
}

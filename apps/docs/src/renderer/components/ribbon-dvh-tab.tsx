import { useEffect, useState, type ReactNode } from 'react'
import { useI18n, type StringKey } from '../i18n/locale'
import { showToast } from './toast-bus'
import { BIG } from './ribbon-tabs'
import {
  IconBatch,
  IconCondition,
  IconConvertTemplate,
  IconExchange,
  IconHistory,
  IconInsertColumnValue,
  IconLinkWorkbook,
  IconProjectData,
  IconRecord,
  IconRefresh,
  IconRelease,
  IconRepeatSection,
  IconScript,
  IconSmartData,
  IconStop,
  IconTemplate,
  IconWorkflow,
  IconWriteBack,
} from './icons'
import {
  docsRecorder,
  onRecorderChange,
  runUiAction,
  startRecording,
  stopRecording,
} from '../dvh-workflow'

/** What a DVH ribbon button asks the Smart Data panel to show (and do). */
export interface DvhPanelRequest {
  readonly section: 'links' | 'fields' | 'tables' | 'template' | 'workflows' | 'project' | 'history'
  readonly action?: 'link' | 'writeback'
  readonly view?: 'blocks' | 'script'
}

/** Opens the Smart Data panel at a section (DvhSmartDataHost listens). */
export function openDvhPanel(request: DvhPanelRequest): void {
  window.dispatchEvent(new CustomEvent<DvhPanelRequest>('dvh:smart-data', { detail: request }))
}

/**
 * The DVH ribbon tab of Docs: every Smart Data feature in one place, grouped
 * as the work goes — data and links, Smart Templates, automation (recorder,
 * workflows, scripts), project data, history.
 */
export function DvhTab({ hasDoc }: { hasDoc: boolean }) {
  const { t } = useI18n()
  const [, setTick] = useState(0)
  useEffect(() => onRecorderChange(() => setTick((n) => n + 1)), [])
  const recording = docsRecorder.recording

  const big = (
    label: StringKey,
    tip: StringKey,
    icon: ReactNode,
    onClick: () => void,
    extra: { pressed?: boolean; disabled?: boolean } = {},
  ) => (
    <button
      className="rb-big"
      disabled={extra.disabled ?? !hasDoc}
      aria-pressed={extra.pressed}
      data-tip={t(tip)}
      onClick={onClick}
    >
      <span className="rb-big-icon">{icon}</span>
      <span>{t(label)}</span>
    </button>
  )
  const small = (label: StringKey, tip: StringKey, icon: ReactNode, onClick: () => void) => (
    <button className="rb-small" disabled={!hasDoc} data-tip={t(tip)} onClick={onClick}>
      {icon} {t(label)}
    </button>
  )

  const updateAll = async () => {
    try {
      const results = (await runUiAction('Link.Update', {})) as {
        updated: { fields: number; tables: number } | null
      }[]
      const count = results.reduce(
        (n, r) => n + (r.updated ? r.updated.fields + r.updated.tables : 0),
        0,
      )
      const missing = results.filter((r) => !r.updated).length
      showToast(t('dvhRbUpdated', { count, missing }))
    } catch (e) {
      showToast(e instanceof Error ? e.message : String(e))
    }
  }

  return (
    <>
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          {big('dvhSmartData', 'dvhSmartDataTip', <IconSmartData size={BIG} />, () =>
            openDvhPanel({ section: 'fields' }),
          )}
          <div className="rb-col">
            {small('dvhRbLink', 'dvhRbLinkTip', <IconLinkWorkbook size={14} />, () =>
              openDvhPanel({ section: 'links', action: 'link' }),
            )}
            {small(
              'dvhRbUpdateAll',
              'dvhRbUpdateAllTip',
              <IconRefresh size={14} />,
              () => void updateAll(),
            )}
            {small('dvhRbWriteBack', 'dvhRbWriteBackTip', <IconWriteBack size={14} />, () =>
              openDvhPanel({ section: 'links', action: 'writeback' }),
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('dvhRbGroupData')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          {big('dvhRbTemplate', 'dvhRbTemplateTip', <IconTemplate size={BIG} />, () =>
            openDvhPanel({ section: 'template' }),
          )}
          <div className="rb-col">
            {small('dvhRbCondition', 'dvhRbConditionTip', <IconCondition size={14} />, () =>
              openDvhPanel({ section: 'template' }),
            )}
            {small('dvhRbRepeat', 'dvhRbRepeatTip', <IconRepeatSection size={14} />, () =>
              openDvhPanel({ section: 'template' }),
            )}
            {small('dvhRbColumn', 'dvhRbColumnTip', <IconInsertColumnValue size={14} />, () =>
              openDvhPanel({ section: 'template' }),
            )}
          </div>
          {big('dvhRbBatch', 'dvhRbBatchTip', <IconBatch size={BIG} />, () =>
            openDvhPanel({ section: 'template' }),
          )}
          {big('dvhRbConvert', 'dvhRbConvertTip', <IconConvertTemplate size={BIG} />, () =>
            openDvhPanel({ section: 'template' }),
          )}
        </div>
        <div className="ribbon-group-label">{t('dvhTemplate')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          {big(
            recording ? 'dvhWfStop' : 'dvhWfRecord',
            'dvhRbRecordTip',
            recording ? <IconStop size={BIG} /> : <IconRecord size={BIG} />,
            () => {
              if (recording) {
                stopRecording()
                openDvhPanel({ section: 'workflows' })
              } else startRecording()
            },
            { pressed: recording },
          )}
          {big('dvhWorkflows', 'dvhRbWorkflowsTip', <IconWorkflow size={BIG} />, () =>
            openDvhPanel({ section: 'workflows', view: 'blocks' }),
          )}
          {big('dvhWfScript', 'dvhRbScriptTip', <IconScript size={BIG} />, () =>
            openDvhPanel({ section: 'workflows', view: 'script' }),
          )}
        </div>
        <div className="ribbon-group-label">{t('dvhRbGroupAutomation')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          {big('dvhProject', 'dvhRbProjectTip', <IconProjectData size={BIG} />, () =>
            openDvhPanel({ section: 'project' }),
          )}
          {big('dvhRbSync', 'dvhProjectSync', <IconExchange size={BIG} />, () =>
            openDvhPanel({ section: 'project' }),
          )}
        </div>
        <div className="ribbon-group-label">{t('dvhProject')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          {big('dvhHistory', 'dvhRbHistoryTip', <IconHistory size={BIG} />, () =>
            openDvhPanel({ section: 'history' }),
          )}
          {big('dvhRbRelease', 'dvhRbReleaseTip', <IconRelease size={BIG} />, () =>
            openDvhPanel({ section: 'history' }),
          )}
        </div>
        <div className="ribbon-group-label">{t('dvhHistory')}</div>
      </div>
    </>
  )
}

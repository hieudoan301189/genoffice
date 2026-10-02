import { useEffect, useRef, useState } from 'react'
import type { Scalar } from '@genoffice/dvh-model'
import {
  describeStep,
  emptyWorkflow,
  insertStep,
  loadWorkflows,
  moveStep,
  parseWorkflow,
  recordsOf,
  removeStep,
  replaceStep,
  stepAt,
  stepSchema,
  walkSteps,
  workflowFileName,
  workflowFromJson,
  workflowToJson,
  wrapStep,
  type Json,
  type PauseDecision,
  type PauseInfo,
  type Step,
  type Workflow,
} from '@genoffice/dvh-workflow'
import { lineOfPath, parseScript, printScript } from '@genoffice/dvh-script'
import { useI18n } from '../i18n/locale'
import { DvhScriptEditor } from './DvhScriptEditor'
import { showToast } from './toast-bus'
import type { DvhDocsState } from '../dvh-smart-data'
import {
  docsRecorder,
  docsUiRegistry,
  onRecorderChange,
  runDocsWorkflow,
  runUiAction,
  setRecordSelection,
  startRecording,
  stopRecording,
} from '../dvh-workflow'

/** The default value of each parameter, as the run form's starting JSON. */
const defaultParams = (workflow: Workflow) =>
  JSON.stringify(
    Object.fromEntries(workflow.params.map((p) => [p.name, p.default ?? null])),
    null,
    2,
  )

/**
 * Workflows of the Smart Data panel (P8): record what the user does through
 * the Action Core, edit the steps as blocks, run them (at once, as a preview
 * or step by step with breakpoints) and keep them in the template or in a
 * project folder.
 */
export function DvhWorkflowSection({
  dvh,
  onChange,
}: {
  dvh: DvhDocsState | null
  onChange: () => void
}) {
  const { t } = useI18n()
  const model = dvh?.model ?? null
  const collections = model?.collections ?? []
  const stored = model ? loadWorkflows(model) : { workflows: [], skipped: [] }
  const [, setTick] = useState(0)
  const [selCollection, setSelCollection] = useState('')
  const [selRecord, setSelRecord] = useState('1')
  const [name, setName] = useState('')
  const [forEach, setForEach] = useState(true)
  const [draft, setDraft] = useState<Workflow | null>(null)
  const [selected, setSelected] = useState<string | null>(null)
  const [editing, setEditing] = useState<string | null>(null)
  const [stepJson, setStepJson] = useState('')
  const [newAction, setNewAction] = useState('')
  const [breakpoints, setBreakpoints] = useState<ReadonlySet<string>>(new Set())
  const [params, setParams] = useState('{}')
  const [dataFrom, setDataFrom] = useState('')
  const [paused, setPaused] = useState<PauseInfo | null>(null)
  /** P9: the draft as blocks or as DVH-Script */
  const [view, setView] = useState<'blocks' | 'script'>('blocks')
  const [scriptText, setScriptText] = useState('')
  const [scriptLines, setScriptLines] = useState<Readonly<Record<string, number>>>({})
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const resume = useRef<((decision: PauseDecision) => void) | null>(null)

  useEffect(() => onRecorderChange(() => setTick((n) => n + 1)), [])

  // the selected record is what the recorder turns into parameters
  const collection = collections.find((c) => c.id === selCollection)
  const recordIndex = Math.max(0, (Number(selRecord) || 1) - 1)
  const record: Record<string, Scalar> | null = collection
    ? (recordsOf(collection)[recordIndex] ?? null)
    : null
  useEffect(() => {
    setRecordSelection(record)
  })
  useEffect(() => () => setRecordSelection(null), [])

  const pick = (workflow: Workflow | null) => {
    setDraft(workflow)
    setSelected(workflow?.id ?? null)
    setEditing(null)
    setBreakpoints(new Set())
    setParams(workflow ? defaultParams(workflow) : '{}')
    setScriptText(workflow ? printScript(workflow) : '')
    setScriptLines({})
    setMessage('')
  }

  const save = async (workflow: Workflow) => {
    try {
      await runUiAction('Workflow.Save', { workflow })
      pick(workflow)
      onChange()
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e))
    }
  }

  const saveRecording = () => {
    const workflow = docsRecorder.toWorkflow({
      name: name.trim() || t('dvhWorkflows'),
      forEach,
      ...(docsUiRegistry() ? { catalog: docsUiRegistry()!.fingerprint() } : {}),
    })
    setName('')
    void save(workflow)
  }

  const edit = (next: Workflow) => {
    setDraft(next)
    setScriptText(printScript(next))
    setEditing(null)
  }

  /** Compiles the script into the draft; null (and the errors shown) when it does not parse. */
  const applyScript = (): { workflow: Workflow; lines: Record<string, number> } | null => {
    const parsed = parseScript(scriptText)
    if (!parsed.ok) {
      setMessage(parsed.errors.map((e) => `${e.line}:${e.column} ${e.message}`).join('\n'))
      return null
    }
    setDraft(parsed.workflow)
    setScriptLines(parsed.lines)
    setMessage('')
    return { workflow: parsed.workflow, lines: { ...parsed.lines } }
  }

  const switchView = (next: 'blocks' | 'script') => {
    if (next === view) return
    if (next === 'blocks' && !applyScript()) return
    if (next === 'script' && draft) setScriptText(printScript(draft))
    setView(next)
  }

  const applyStep = (path: string) => {
    if (!draft) return
    try {
      const step = stepSchema.parse(JSON.parse(stepJson)) as Step
      edit(replaceStep(draft, path, step))
      setMessage('')
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e))
    }
  }

  const addAction = () => {
    if (!draft || !newAction) return
    edit(
      insertStep(draft, String(draft.steps.length), {
        kind: 'action',
        action: newAction,
        args: {},
      }),
    )
  }

  const decide = (decision: PauseDecision) => {
    resume.current?.(decision)
    resume.current = null
    setPaused(null)
  }

  const run = async (mode: 'run' | 'preview' | 'debug') => {
    if (!draft || busy) return
    let values: Record<string, Json>
    try {
      values = JSON.parse(params) as Record<string, Json>
      const coll = collections.find((c) => c.id === dataFrom)
      const list = draft.params.find((p) => p.type === 'list')
      if (coll && list) values = { ...values, [list.name]: recordsOf(coll) }
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e))
      return
    }
    // a script runs as written: compile it first
    const compiled = view === 'script' ? applyScript() : null
    if (view === 'script' && !compiled) return
    setBusy(true)
    setMessage('')
    try {
      const workflow = compiled?.workflow ?? parseWorkflow(draft)
      const result = await runDocsWorkflow(workflow, {
        params: values,
        dryRun: mode === 'preview',
        stepMode: mode === 'debug',
        breakpoints,
        onPause: (pause) =>
          new Promise<PauseDecision>((resolve) => {
            resume.current = resolve
            setPaused(pause)
          }),
        confirm: (action, preview) =>
          window.confirm(
            t('dvhWfConfirm', {
              action,
              count: preview.objects,
              summary: preview.summary.join('\n'),
            }),
          ),
      })
      const lines = result.log
        .filter((e) => e.kind !== 'pause')
        .slice(-12)
        .map((e) => `${e.path || '·'} ${e.kind}: ${e.message}`)
      if (result.ok) {
        const actions = mode === 'preview' ? result.previews.length : result.changeSets.length
        showToast(t('dvhWfDone', { steps: result.steps, actions }))
      } else {
        const line = compiled ? lineOfPath(compiled.lines, result.failedAt) : null
        lines.push(
          t('dvhWfFailed', {
            path: line
              ? `${result.failedAt ?? '·'} (${t('dvhWfLine', { line })})`
              : (result.failedAt ?? '·'),
            error: result.error ?? '',
          }),
        )
      }
      setMessage(lines.join('\n'))
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e))
    } finally {
      resume.current = null
      setPaused(null)
      setBusy(false)
      onChange()
    }
  }

  const exportFile = async () => {
    if (!draft) return
    const path = await window.desktop.dvhWorkflowExport?.(
      workflowToJson(draft),
      workflowFileName(draft),
    )
    if (path) showToast(path)
  }

  const importFile = async () => {
    const picked = await window.desktop.dvhWorkflowImport?.()
    if (!picked) return
    try {
      await save(workflowFromJson(picked.text))
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e))
    }
  }

  const remove = async () => {
    if (!selected) return
    try {
      await runUiAction('Workflow.Delete', { workflow: selected })
      pick(null)
      onChange()
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e))
    }
  }

  const recording = docsRecorder.recording
  // the whole catalog for the script editor; the block editor offers actions that change something
  const catalog = docsUiRegistry()?.catalog() ?? []
  const writeActions = catalog.filter((a) => a.effect !== 'read')

  return (
    <details className="dvh-docs-panel-section" data-section="workflows">
      <summary>{t('dvhWorkflows')}</summary>

      <div className="dvh-docs-panel-new" data-status="recorder">
        <select
          aria-label={t('dvhWfRecordFrom')}
          value={selCollection}
          onChange={(event) => setSelCollection(event.target.value)}
        >
          <option value="">{t('dvhWfNoRecord')}</option>
          {collections.map((c) => (
            <option key={c.id} value={c.id}>
              {t('dvhWfRecordFrom')}: {c.name}
            </option>
          ))}
        </select>
        <input
          type="number"
          min={1}
          value={selRecord}
          aria-label={t('dvhWfRecordFrom')}
          disabled={!collection}
          onChange={(event) => setSelRecord(event.target.value)}
        />
        {recording ? (
          <button type="button" onClick={stopRecording}>
            {t('dvhWfStop')}
          </button>
        ) : (
          <button type="button" onClick={startRecording}>
            {t('dvhWfRecord')}
          </button>
        )}
      </div>
      {recording || docsRecorder.steps().length > 0 ? (
        <div className="dvh-docs-panel-workflow" data-status={recording ? 'recording' : 'recorded'}>
          <p className="dvh-docs-panel-muted">
            {t('dvhWfRecording', { count: docsRecorder.steps().length })}
          </p>
          <ol className="dvh-docs-panel-steps">
            {docsRecorder.steps().map((s, i) => (
              <li key={i}>{`${s.action}(${JSON.stringify(s.input)})`}</li>
            ))}
          </ol>
          {docsRecorder.unrecorded().map((command) => (
            <p key={command} className="dvh-docs-panel-warn">
              {t('dvhWfUnrecordable', { command })}
            </p>
          ))}
          {!recording && docsRecorder.steps().length > 0 ? (
            <div className="dvh-docs-panel-new">
              <input
                value={name}
                placeholder={t('dvhWfName')}
                aria-label={t('dvhWfName')}
                onChange={(event) => setName(event.target.value)}
              />
              <label>
                <input
                  type="checkbox"
                  checked={forEach}
                  onChange={(event) => setForEach(event.target.checked)}
                />
                {t('dvhWfForEach')}
              </label>
              <button type="button" onClick={saveRecording}>
                {t('dvhWfSaveRecording')}
              </button>
            </div>
          ) : null}
        </div>
      ) : null}

      <div className="dvh-docs-panel-new">
        <select
          aria-label={t('dvhWorkflows')}
          value={selected ?? ''}
          onChange={(event) =>
            pick(stored.workflows.find((w) => w.id === event.target.value) ?? null)
          }
        >
          <option value="">{stored.workflows.length ? '—' : t('dvhWfNone')}</option>
          {stored.workflows.map((w) => (
            <option key={w.id} value={w.id}>
              {w.name}
            </option>
          ))}
        </select>
        <button type="button" onClick={() => pick(emptyWorkflow(t('dvhWorkflows')))}>
          +
        </button>
        <button type="button" onClick={() => void importFile()}>
          {t('dvhWfImport')}
        </button>
      </div>
      {stored.skipped.length > 0 ? (
        <p className="dvh-docs-panel-warn">{stored.skipped.join(', ')}</p>
      ) : null}

      {draft ? (
        <div className="dvh-docs-panel-workflow" data-workflow-id={draft.id}>
          <input
            value={draft.name}
            aria-label={t('dvhWfName')}
            onChange={(event) => setDraft({ ...draft, name: event.target.value })}
          />
          <div className="dvh-docs-panel-link-actions" role="tablist">
            <button
              type="button"
              role="tab"
              aria-selected={view === 'blocks'}
              onClick={() => switchView('blocks')}
            >
              {t('dvhWfBlocks')}
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={view === 'script'}
              onClick={() => switchView('script')}
            >
              {t('dvhWfScript')}
            </button>
          </div>
          {view === 'script' ? (
            <>
              <DvhScriptEditor
                value={scriptText}
                onChange={setScriptText}
                catalog={catalog}
                label={t('dvhWfScript')}
                pausedLine={paused ? lineOfPath(scriptLines, paused.path) : null}
              />
              <div className="dvh-docs-panel-link-actions">
                <button type="button" onClick={() => void applyScript()}>
                  {t('dvhWfApplyScript')}
                </button>
              </div>
            </>
          ) : (
            <>
              <ol className="dvh-docs-panel-blocks">
                {walkSteps(draft.steps).map(({ path, step, depth }) => (
                  <li
                    key={path}
                    data-path={path}
                    data-kind={step.kind}
                    data-current={paused?.path === path ? 'true' : undefined}
                    style={{ marginInlineStart: `${depth * 12}px` }}
                  >
                    <label title={t('dvhWfBreakpoint')}>
                      <input
                        type="checkbox"
                        aria-label={t('dvhWfBreakpoint')}
                        checked={breakpoints.has(path)}
                        onChange={(event) => {
                          const next = new Set(breakpoints)
                          if (event.target.checked) next.add(path)
                          else next.delete(path)
                          setBreakpoints(next)
                        }}
                      />
                    </label>
                    <code>{describeStep(step)}</code>
                    <span className="dvh-docs-panel-link-actions">
                      <button
                        type="button"
                        aria-label={t('dvhWfMoveUp')}
                        onClick={() => edit(moveStep(draft, path, -1))}
                      >
                        ↑
                      </button>
                      <button
                        type="button"
                        aria-label={t('dvhWfMoveDown')}
                        onClick={() => edit(moveStep(draft, path, 1))}
                      >
                        ↓
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setEditing(path)
                          setStepJson(JSON.stringify(stepAt(draft, path), null, 2))
                        }}
                      >
                        {t('dvhWfEdit')}
                      </button>
                      <button
                        type="button"
                        onClick={() => edit(wrapStep(draft, path, 'transaction'))}
                      >
                        {t('dvhWfWrapTx')}
                      </button>
                      <button type="button" onClick={() => edit(wrapStep(draft, path, 'try'))}>
                        {t('dvhWfWrapTry')}
                      </button>
                      <button
                        type="button"
                        aria-label={t('dvhWfRemove')}
                        onClick={() => edit(removeStep(draft, path))}
                      >
                        ×
                      </button>
                    </span>
                    {editing === path ? (
                      <div className="dvh-docs-panel-new">
                        <textarea
                          value={stepJson}
                          rows={6}
                          aria-label={t('dvhWfEdit')}
                          onChange={(event) => setStepJson(event.target.value)}
                        />
                        <button type="button" onClick={() => applyStep(path)}>
                          {t('dvhWfApply')}
                        </button>
                      </div>
                    ) : null}
                  </li>
                ))}
              </ol>
              <div className="dvh-docs-panel-new">
                <select
                  aria-label={t('dvhWfAddStep')}
                  value={newAction}
                  onChange={(event) => setNewAction(event.target.value)}
                >
                  <option value="">{t('dvhWfAddStep')}</option>
                  {writeActions.map((a) => (
                    <option key={a.name} value={a.name}>
                      {a.name}
                    </option>
                  ))}
                </select>
                <button type="button" disabled={!newAction} onClick={addAction}>
                  +
                </button>
              </div>
            </>
          )}

          <textarea
            value={params}
            rows={4}
            aria-label={t('dvhWfParams')}
            placeholder={t('dvhWfParams')}
            onChange={(event) => setParams(event.target.value)}
          />
          <div className="dvh-docs-panel-new">
            <select
              aria-label={t('dvhWfDataFrom')}
              value={dataFrom}
              onChange={(event) => setDataFrom(event.target.value)}
            >
              <option value="">{t('dvhWfParams')}</option>
              {collections.map((c) => (
                <option key={c.id} value={c.id}>
                  {t('dvhWfDataFrom')}: {c.name}
                </option>
              ))}
            </select>
            <button type="button" disabled={busy} onClick={() => void run('run')}>
              {t('dvhWfRun')}
            </button>
            <button type="button" disabled={busy} onClick={() => void run('preview')}>
              {t('dvhWfPreview')}
            </button>
            <button type="button" disabled={busy} onClick={() => void run('debug')}>
              {t('dvhWfDebug')}
            </button>
          </div>
          {paused ? (
            <div className="dvh-docs-panel-conflict" data-status="paused">
              <strong>{t('dvhWfPaused', { path: paused.path })}</strong>
              <code className="dvh-docs-panel-pre">{JSON.stringify(paused.variables)}</code>
              <div className="dvh-docs-panel-link-actions">
                <button type="button" onClick={() => decide('continue')}>
                  {t('dvhWfContinue')}
                </button>
                <button type="button" onClick={() => decide('step')}>
                  {t('dvhWfStep')}
                </button>
                <button type="button" onClick={() => decide('stop')}>
                  {t('dvhWfHalt')}
                </button>
              </div>
            </div>
          ) : null}
          <div className="dvh-docs-panel-link-actions">
            <button type="button" disabled={busy} onClick={() => void save(draft)}>
              {t('dvhWfSave')}
            </button>
            <button type="button" disabled={busy} onClick={() => void exportFile()}>
              {t('dvhWfExport')}
            </button>
            <button type="button" disabled={busy || !selected} onClick={() => void remove()}>
              {t('dvhWfDelete')}
            </button>
          </div>
        </div>
      ) : null}
      {message ? <p className="dvh-docs-panel-muted dvh-docs-panel-pre">{message}</p> : null}
    </details>
  )
}

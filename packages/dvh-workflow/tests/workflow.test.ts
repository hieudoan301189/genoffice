/**
 * P8: the workflow engine (control flow, variables, transactions, rollback,
 * step-by-step runs, limits) and the recorder over the Action Core event
 * stream. Acceptance: a recording does not depend on which UI ran the
 * actions, and a recorded workflow runs again with another data set.
 */
import { describe, expect, it } from 'vitest'
import {
  ActionRegistry,
  localParticipant,
  z,
  type ActionGroup,
  type ActionEffect,
  type SagaParticipant,
} from '@genoffice/dvh-actions'
import { emptyModel, parseModelXml, serializeModelXml, type Scalar } from '@genoffice/dvh-model'
import {
  WorkflowRecorder,
  describeStep,
  emptyWorkflow,
  insertStep,
  loadWorkflows,
  moveStep,
  parseWorkflow,
  removeStep,
  removeWorkflow,
  runWorkflow,
  stepAt,
  storeWorkflow,
  walkSteps,
  workflowFileName,
  workflowFromJson,
  workflowIssues,
  workflowToJson,
  wrapStep,
  type Workflow,
} from '../src/index'

/** An in-memory document with a handful of actions. */
function fakeDoc(docId: string) {
  let state = { fields: {} as Record<string, Scalar>, rows: [] as string[] }
  const registry = new ActionRegistry()
  const add = (
    name: string,
    group: ActionGroup,
    effect: ActionEffect,
    input: z.ZodType,
    run: (input: never) => unknown,
    objects = 1,
  ) =>
    registry.register({
      name,
      group,
      summary: name,
      effect,
      input,
      preview: () => ({ objects, files: [], summary: [name], warnings: [] }),
      execute: (i, ctx) => {
        const out = run(i as never)
        if (effect !== 'read')
          ctx.emit([{ op: 'set', target: { kind: 'field', id: name } }] as never)
        return out
      },
    })
  add(
    'Data.SetField',
    'Data',
    'write',
    z.object({ field: z.string(), value: z.union([z.string(), z.number(), z.null()]) }).strict(),
    (i: { field: string; value: Scalar }) => {
      state.fields[i.field] = i.value
    },
  )
  add(
    'Data.GetField',
    'Data',
    'read',
    z.object({ field: z.string() }).strict(),
    (i: { field: string }) => state.fields[i.field] ?? null,
  )
  add(
    'Data.AddRow',
    'Data',
    'write',
    z.object({ text: z.string() }).strict(),
    (i: { text: string }) => {
      if (i.text === 'boom') throw new Error('disk full')
      state.rows.push(i.text)
      return state.rows.length
    },
  )
  add('Data.DeleteAll', 'Data', 'destructive', z.object({}).strict(), () => {
    state.rows = []
  })
  add('UI.Select', 'UI', 'read', z.object({ field: z.string() }).strict(), () => null)
  add('Document.Undo', 'Document', 'write', z.object({}).strict(), () => null)
  add('Document.Redo', 'Document', 'write', z.object({}).strict(), () => null)
  const participant: SagaParticipant = localParticipant({
    docId,
    registry,
    checkpoint: () => structuredClone(state),
    restore: (cp) => {
      state = structuredClone(cp as typeof state)
    },
  })
  return { registry, participant, state: () => state }
}

const run = (workflow: Workflow, doc: ReturnType<typeof fakeDoc>, extra = {}) =>
  runWorkflow(workflow, {
    participants: new Map([['doc_a', doc.participant]]),
    docId: 'doc_a',
    ...extra,
  })

const wf = (steps: Workflow['steps'], params: Workflow['params'] = []): Workflow =>
  parseWorkflow({ ...emptyWorkflow('test', 'wf_test'), params, steps })

describe('workflow engine', () => {
  it('runs ForEach, If, Set, assignments and expressions as one transaction', async () => {
    const doc = fakeDoc('doc_a')
    const workflow = wf(
      [
        { kind: 'set', variable: 'total', value: { value: 0 } },
        {
          kind: 'forEach',
          variable: 'item',
          items: { expr: 'items' },
          where: 'item.qty > 0',
          steps: [
            {
              kind: 'action',
              action: 'Data.AddRow',
              args: { text: { expr: 'PAD(INDEX, 2) & "-" & UPPER(item.name)' } },
              assign: 'count',
            },
            { kind: 'set', variable: 'total', value: { expr: 'total + item.qty' } },
          ],
        },
        {
          kind: 'if',
          branches: [
            { when: 'total > 100', steps: [{ kind: 'log', message: { value: 'big' } }] },
            {
              when: 'total > 5',
              steps: [
                {
                  kind: 'action',
                  action: 'Data.SetField',
                  args: { field: { value: 'Total' }, value: { expr: 'total' } },
                },
              ],
            },
          ],
          else: [{ kind: 'log', message: { value: 'small' } }],
        },
        { kind: 'log', message: { expr: '"rows: " & items.Count' } },
      ],
      [{ name: 'items', type: 'list' }],
    )
    const items = [
      { name: 'móng', qty: 3 },
      { name: 'cột', qty: 0 },
      { name: 'dầm', qty: 4 },
    ]
    const result = await run(workflow, doc, { params: { items } })
    expect(result.ok).toBe(true)
    expect(doc.state().rows).toEqual(['01-MÓNG', '03-DẦM'])
    expect(doc.state().fields.Total).toBe(7)
    expect(result.variables).toMatchObject({ total: 7, count: 2 })
    expect(result.log.filter((e) => e.kind === 'log').map((e) => e.message)).toEqual(['rows: 3'])
    // every change set of the run carries the run's txId
    expect(result.changeSets).toHaveLength(3)
    expect(new Set(result.changeSets.map((c) => c.txId))).toEqual(new Set([result.txId]))
    expect(result.changeSets[0]!.source).toBe('workflow')
  })

  it('a dry run previews every action and changes nothing', async () => {
    const doc = fakeDoc('doc_a')
    const result = await run(
      wf([
        { kind: 'action', action: 'Data.AddRow', args: { text: { value: 'a' } } },
        {
          kind: 'action',
          action: 'Data.SetField',
          args: { field: { value: 'X' }, value: { value: 1 } },
        },
      ]),
      doc,
      { dryRun: true },
    )
    expect(result.ok).toBe(true)
    expect(result.previews.map((p) => p.action)).toEqual(['Data.AddRow', 'Data.SetField'])
    expect(doc.state().rows).toEqual([])
  })

  it('rolls back a failed Transaction only, lets Try carry on, and rolls back the run on an uncaught error', async () => {
    const doc = fakeDoc('doc_a')
    const add = (text: string) => ({
      kind: 'action' as const,
      action: 'Data.AddRow',
      args: { text: { value: text } },
    })
    const caught = await run(
      wf([
        add('kept'),
        {
          kind: 'try',
          steps: [{ kind: 'transaction', steps: [add('undone'), add('boom')] }],
          catch: [
            { kind: 'log', message: { expr: '"caught at " & error.step & ": " & error.message' } },
          ],
        },
        add('after'),
      ]),
      doc,
    )
    expect(caught.ok).toBe(true)
    expect(doc.state().rows).toEqual(['kept', 'after'])
    expect(caught.log.find((e) => e.kind === 'log')?.message).toBe(
      'caught at 1.body.0.body.1: disk full',
    )
    expect(caught.log.some((e) => e.kind === 'rollback')).toBe(true)

    const failed = await run(wf([add('one'), add('boom'), add('never')]), doc)
    expect(failed.ok).toBe(false)
    expect(failed.failedAt).toBe('1')
    expect(failed.error).toBe('disk full')
    expect(failed.rolledBack).toEqual(['doc_a'])
    // the whole run is undone: the rows are what they were before it
    expect(doc.state().rows).toEqual(['kept', 'after'])
  })

  it('pauses step by step and at breakpoints; stop halts and rolls back', async () => {
    const doc = fakeDoc('doc_a')
    const workflow = wf([
      { kind: 'action', action: 'Data.AddRow', args: { text: { value: 'a' } } },
      {
        kind: 'forEach',
        variable: 'x',
        items: { value: [1, 2] },
        steps: [{ kind: 'set', variable: 'last', value: { expr: 'x' } }],
      },
      { kind: 'action', action: 'Data.AddRow', args: { text: { value: 'b' } } },
    ])
    const paused: string[] = []
    const stepped = await run(workflow, doc, {
      stepMode: true,
      onPause: (p: { path: string }) => {
        paused.push(p.path)
        return 'step'
      },
    })
    expect(stepped.ok).toBe(true)
    expect(paused).toEqual(['0', '1', '1.body.0', '1.body.0', '2'])
    expect(walkSteps(workflow.steps).map((s) => s.path)).toEqual(['0', '1', '1.body.0', '2'])

    const seen: Record<string, unknown>[] = []
    const stopped = await run(workflow, doc, {
      breakpoints: new Set(['2']),
      onPause: (p: { variables: Record<string, unknown> }) => {
        seen.push(p.variables)
        return 'stop'
      },
    })
    expect(stopped.ok).toBe(false)
    expect(stopped.stopped).toBe(true)
    expect(seen).toEqual([{ last: 2 }])
    expect(doc.state().rows).toEqual(['a', 'b'])
  })

  it('limits steps (not catchable by Try), checks permissions and parameters', async () => {
    const doc = fakeDoc('doc_a')
    const endless = await run(
      wf(
        [
          {
            kind: 'try',
            steps: [
              {
                kind: 'forEach',
                variable: 'i',
                items: { expr: 'big' },
                steps: [{ kind: 'set', variable: 'n', value: { expr: 'i' } }],
              },
            ],
            catch: [],
          },
        ],
        [{ name: 'big', type: 'list' }],
      ),
      doc,
      { params: { big: Array.from({ length: 500 }, (_, i) => i) }, maxSteps: 100 },
    )
    expect(endless.ok).toBe(false)
    expect(endless.error).toContain('exceeded 100 steps')

    const destructive = await run(wf([{ kind: 'action', action: 'Data.DeleteAll', args: {} }]), doc)
    expect(destructive.ok).toBe(false)
    expect(destructive.error).toContain('may not run destructive')

    const missing = await run(wf([], [{ name: 'n', type: 'number' }]), doc)
    expect(missing.error).toBe('parameter n has no value')
    const coerced = await run(
      wf(
        [{ kind: 'set', variable: 'm', value: { expr: 'n * 2' } }],
        [{ name: 'n', type: 'number' }],
      ),
      doc,
      { params: { n: '21' } },
    )
    expect(coerced.variables.m).toBe(42)
  })

  it('rejects malformed workflows and unknown actions', () => {
    expect(() =>
      parseWorkflow({
        ...emptyWorkflow('x'),
        steps: [{ kind: 'set', variable: '1x', value: { value: 1 } }],
      }),
    ).toThrow(/steps\.0\.variable/)
    expect(() =>
      parseWorkflow({ ...emptyWorkflow('x'), steps: [{ kind: 'log', message: { expr: 'a +' } }] }),
    ).toThrow(/bad expression/)
    const doc = fakeDoc('doc_a')
    expect(
      workflowIssues(wf([{ kind: 'action', action: 'File.DeleteAll', args: {} }]), doc.registry),
    ).toEqual([{ path: '0', message: 'File.DeleteAll is not in the catalog' }])
  })
})

describe('recorder', () => {
  /** two "layouts": a ribbon and a side panel, both going through the Action Core */
  const ribbon = (doc: ReturnType<typeof fakeDoc>) => ({
    set: (field: string, value: string) => doc.registry.run('Data.SetField', { field, value }, ui),
    select: (field: string) => doc.registry.run('UI.Select', { field }, ui),
    add: (text: string) => doc.registry.run('Data.AddRow', { text }, ui),
    undo: () => doc.registry.run('Document.Undo', {}, ui),
    redo: () => doc.registry.run('Document.Redo', {}, ui),
  })
  const panel = (doc: ReturnType<typeof fakeDoc>) => {
    // a different layout calls the same actions in its own way
    const call = (name: string, input: unknown) => doc.registry.run(name, input, ui)
    return {
      set: (field: string, value: string) => call('Data.SetField', { value, field }),
      select: (field: string) => call('UI.Select', { field }),
      add: (text: string) => call('Data.AddRow', { text }),
      undo: () => call('Document.Undo', {}),
      redo: () => call('Document.Redo', {}),
    }
  }
  const ui = {
    caller: 'ui' as const,
    docId: 'doc_a',
    permissions: new Set<ActionEffect>(['read', 'write']),
  }

  async function session(layout: typeof ribbon) {
    const doc = fakeDoc('doc_a')
    let selected: Record<string, Scalar> = { code: 'HM-01', name: 'Móng M1' }
    const recorder = new WorkflowRecorder({ context: () => ({ record: { values: selected } }) })
    recorder.attach(doc.registry)
    const ui = layout(doc)
    await ui.add('not recorded: before start')
    recorder.start()
    await ui.select('Project.Code')
    await ui.set('Project.Code', 'H')
    await ui.set('Project.Code', 'HM')
    await ui.set('Project.Code', 'HM-01')
    await ui.set('Project.Title', 'Biên bản Móng M1')
    await ui.add('typo')
    await ui.undo()
    await ui.add('Móng M1')
    await ui.undo()
    await ui.redo()
    selected = { code: 'HM-01', name: 'Móng M1' }
    recorder.unrecordable('Link workbook')
    recorder.stop()
    await ui.add('not recorded: after stop')
    // agent and workflow runs are never recorded
    await doc.registry.run('Data.AddRow', { text: 'ai' }, { ...ui_ai })
    return { recorder, doc }
  }
  const ui_ai = {
    caller: 'ai' as const,
    docId: 'doc_a',
    permissions: new Set<ActionEffect>(['write']),
  }

  it('normalises the event stream: merges setters, drops selections, cancels undo/redo pairs', async () => {
    const { recorder } = await session(ribbon)
    expect(recorder.steps().map((s) => [s.action, s.input])).toEqual([
      ['Data.SetField', { field: 'Project.Code', value: 'HM-01' }],
      ['Data.SetField', { field: 'Project.Title', value: 'Biên bản Móng M1' }],
      ['Data.AddRow', { text: 'Móng M1' }],
    ])
    expect(recorder.warnings()).toEqual(['"Link workbook" cannot be recorded'])
  })

  it('acceptance: the same work recorded from another UI layout gives the same workflow', async () => {
    const a = (await session(ribbon)).recorder.toWorkflow({ name: 'Biên bản', id: 'wf_1' })
    const b = (await session(panel)).recorder.toWorkflow({ name: 'Biên bản', id: 'wf_1' })
    expect(workflowToJson(b)).toBe(workflowToJson(a))
    expect(a.steps.map(describeStep)).toEqual([
      'Data.SetField(field:="Project.Code", value:=record.code)',
      'Data.SetField(field:="Project.Title", value:="Biên bản Móng M1")',
      'Data.AddRow(text:=record.name)',
    ])
    expect(a.params).toEqual([
      { name: 'record', type: 'record', default: { code: 'HM-01', name: 'Móng M1' } },
    ])
  })

  it('acceptance: a recorded workflow runs again with another data set', async () => {
    const { recorder } = await session(ribbon)
    const workflow = recorder.toWorkflow({ name: 'Từng hạng mục', forEach: true })
    expect(workflow.steps[0]).toMatchObject({ kind: 'forEach', variable: 'record' })
    const fresh = fakeDoc('doc_a')
    const result = await run(workflow, fresh, {
      params: {
        records: [
          { code: 'HM-07', name: 'Cột C7' },
          { code: 'HM-08', name: 'Dầm D8' },
        ],
      },
    })
    expect(result.ok).toBe(true)
    expect(fresh.state().rows).toEqual(['Cột C7', 'Dầm D8'])
    expect(fresh.state().fields).toEqual({
      'Project.Code': 'HM-08',
      'Project.Title': 'Biên bản Móng M1',
    })
    // with no data it replays the recording itself
    const replay = fakeDoc('doc_a')
    expect((await run(workflow, replay)).ok).toBe(true)
    expect(replay.state().rows).toEqual(['Móng M1'])
  })
})

describe('storage and block edits', () => {
  const sample = wf([
    { kind: 'action', action: 'Data.AddRow', args: { text: { value: 'a' } } },
    { kind: 'action', action: 'Data.AddRow', args: { text: { value: 'b' } } },
  ])

  it('stores workflows in the model part and in project files', () => {
    const model = storeWorkflow(emptyModel('doc_aaaaaaaaaaaaaaaa'), sample)
    const back = parseModelXml(serializeModelXml(model))
    expect(loadWorkflows(back).workflows).toEqual([sample])
    // a workflow from a newer format is skipped, not fatal
    const odd = {
      ...back,
      workflows: [...back.workflows!, { id: 'wf_new', name: 'v2', version: 2 }],
    }
    expect(loadWorkflows(odd)).toMatchObject({ skipped: ['v2'] })
    expect(removeWorkflow(model, sample.id).workflows).toBeUndefined()
    expect(serializeModelXml(removeWorkflow(model, sample.id))).toBe(
      serializeModelXml(emptyModel('doc_aaaaaaaaaaaaaaaa')),
    )
    expect(workflowFromJson(workflowToJson(sample))).toEqual(sample)
    expect(workflowFileName({ id: 'wf_x', name: 'Biên bản: nghiệm thu?' })).toBe(
      'Biên bản nghiệm thu.dvhflow.json',
    )
  })

  it('edits blocks by path without mutating the original', () => {
    const wrapped = wrapStep(sample, '1', 'transaction')
    expect(stepAt(wrapped, '1.body.0')).toEqual(sample.steps[1])
    const inserted = insertStep(wrapped, '1.body.1', { kind: 'log', message: { value: 'x' } })
    expect(walkSteps(inserted.steps).map((s) => s.path)).toEqual(['0', '1', '1.body.0', '1.body.1'])
    const moved = moveStep(inserted, '1.body.1', -1)
    expect(stepAt(moved, '1.body.0')).toEqual({ kind: 'log', message: { value: 'x' } })
    expect(removeStep(moved, '0').steps).toHaveLength(1)
    expect(sample.steps).toHaveLength(2)
    expect(() => stepAt(sample, '0.body.0')).toThrow(/bad step path/)
  })
})

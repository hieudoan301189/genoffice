/**
 * P9 acceptance (Script): a script round-trips without losing any action or
 * argument — `parse(print(w)) == w` for workflows in canonical form, checked
 * on hand-written cases and on a few hundred generated workflows.
 */
import { describe, expect, it } from 'vitest'
import {
  parseWorkflow,
  type Arg,
  type Json,
  type Step,
  type Workflow,
} from '@genoffice/dvh-workflow'
import { canonicalWorkflow, compileScript, parseScript, printScript } from '../src/index'

const SAMPLE = `' Biên bản nghiệm thu: one per work item
Workflow "Biên bản nghiệm thu"
Id "wf_bb"
Param records As List = Json([{"code": "HM-01", "qty": 3}]) Description "the work items"
Param title As Text = "BB ""số"" 1"

total = 0
For Each row In records Where row.qty > 0
    n = Data.AddRow(text:="Hạng mục " & row.code) ' the visible row
    total = total + row.qty
Next
If total > 100 Then
    Log "big"
ElseIf total > 0 Then
    Data.SetField(field:="Total", value:=total) On "doc_report"
Else
    Log "nothing"
End If
Try
    Transaction
        Data.DeleteAll()
    End Transaction
Catch err
    Log "failed: " & err.message
End Try
`

describe('DVH-Script parse and print', () => {
  it('compiles the sample into the workflow model', () => {
    const { workflow, lines } = compileScript(SAMPLE)
    expect(workflow.name).toBe('Biên bản nghiệm thu')
    expect(workflow.params).toEqual([
      {
        name: 'records',
        type: 'list',
        default: [{ code: 'HM-01', qty: 3 }],
        description: 'the work items',
      },
      { name: 'title', type: 'text', default: 'BB "số" 1' },
    ])
    expect(workflow.steps[1]).toEqual({
      kind: 'forEach',
      variable: 'row',
      items: { expr: 'records' },
      where: 'row.qty > 0',
      steps: [
        {
          kind: 'action',
          action: 'Data.AddRow',
          args: { text: { expr: '"Hạng mục " & row.code' } },
          assign: 'n',
        },
        { kind: 'set', variable: 'total', value: { expr: 'total + row.qty' } },
      ],
    })
    expect(workflow.steps[2]).toMatchObject({
      kind: 'if',
      branches: [
        { when: 'total > 100' },
        {
          when: 'total > 0',
          steps: [
            {
              action: 'Data.SetField',
              args: { field: { value: 'Total' }, value: { expr: 'total' } },
              doc: { value: 'doc_report' },
            },
          ],
        },
      ],
      else: [{ kind: 'log', message: { value: 'nothing' } }],
    })
    expect(workflow.steps[3]).toMatchObject({ kind: 'try', errorVariable: 'err' })
    expect(lines).toMatchObject({ '1': 8, '1.body.0': 9, '2.then1.0': 15, '3.body.0.body.0': 21 })
    // print, parse: the same model
    expect(compileScript(printScript(workflow)).workflow).toEqual(workflow)
  })

  it('reports errors with their line and column', () => {
    const cases: [string, number, RegExp][] = [
      ['x = 1\nIf x > 1\n', 2, /If <condition> Then/],
      ['For Each r In rows\nLog r\n', 1, /missing Next/],
      ['Data.SetField(field "a")\n', 1, /name:=value/],
      ['x = (1 +\n', 1, /bad expression/],
      ['Next\n', 1, /Next without For Each/],
      ['Log "a"\nParam p As Text\n', 2, /cannot read/],
      ['Param p As Date\n', 1, /unknown type/],
      ['x = Json([1,)\n', 1, /bad Json/],
      ['Try\nLog 1\nEnd Try\n', 3, /needs a Catch/],
    ]
    for (const [source, line, message] of cases) {
      const outcome = parseScript(source)
      expect(outcome.ok, source).toBe(false)
      if (!outcome.ok) {
        expect(outcome.errors[0]!.message, source).toMatch(message)
        expect(outcome.errors[0]!.line, source).toBe(line)
      }
    }
    const col = parseScript('    Data.SetField(field:=1 +)\n')
    // the column of the faulty value, after "field:="
    expect(col.ok || col.errors[0]!.column).toBe(26)
  })

  it('canonical form: same meaning, script notation', () => {
    const w = parseWorkflow({
      format: 'dvh-workflow',
      version: 1,
      id: 'wf_c',
      name: 'c',
      params: [],
      steps: [
        { kind: 'set', variable: 'a', value: { expr: " 'x' & 'it''s' " } },
        { kind: 'set', variable: 'b', value: { expr: '42' } },
        { kind: 'log', message: { expr: '"plain"' } },
      ],
    })
    const c = canonicalWorkflow(w)
    expect(c.steps).toEqual([
      { kind: 'set', variable: 'a', value: { expr: `"x" & "it's"` } },
      { kind: 'set', variable: 'b', value: { value: 42 } },
      { kind: 'log', message: { value: 'plain' } },
    ])
    expect(compileScript(printScript(w)).workflow).toEqual(c)
  })

  it('round-trips generated workflows without losing actions or arguments', () => {
    let seed = 20261002
    const rnd = () => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff
      return seed / 0x7fffffff
    }
    const pick = <T>(list: readonly T[]): T => list[Math.floor(rnd() * list.length)]!
    const names = ['a', 'b', 'row', 'tổng', 'Else', 'Next', 'If', 'Log', 'x_1']
    const texts = [
      '',
      'plain',
      'with "quotes"',
      "it's",
      'line\nbreak',
      'tab\t',
      'Then',
      ' Where ',
      '\\back',
      'Đường số 7',
      "' not a comment",
      'Json(1)',
    ]
    const scalar = (): Json =>
      pick<() => Json>([
        () => pick(texts),
        () => Math.round(rnd() * 2000 - 1000) / pick([1, 10, 100]) || 1,
        () => rnd() > 0.5,
        () => null,
        () => 1e21,
      ])()
    const value = (depth = 0): Json =>
      depth < 2 && rnd() < 0.25
        ? rnd() < 0.5
          ? [value(depth + 1), value(depth + 1)]
          : { [pick(texts)]: value(depth + 1), k: value(depth + 1) }
        : scalar()
    const exprs = [
      'a + 1',
      'row.qty * 2',
      '"t" & a',
      'IF(a > 1, "x", "y")',
      '[Mã HM] & "-"',
      'NOT b',
      'LEN(row.name) >= 3',
      'records',
      'a.Count',
    ]
    const arg = (): Arg => (rnd() < 0.5 ? { value: value() } : { expr: pick(exprs) })
    const steps = (depth: number): Step[] =>
      Array.from({ length: Math.floor(rnd() * 4) }, (): Step => {
        const kind =
          depth > 2
            ? pick(['action', 'set', 'log'] as const)
            : pick(['action', 'set', 'log', 'if', 'forEach', 'try', 'transaction'] as const)
        switch (kind) {
          case 'action':
            return {
              kind: 'action',
              action: pick([
                'Data.SetField',
                'Table.Refresh',
                'Document.InsertParagraph',
                'Workflow.Run',
              ]),
              args: Object.fromEntries(
                Array.from({ length: Math.floor(rnd() * 3) }, (_, i) => [
                  `${pick(['field', 'value', 'text'])}${i}`,
                  arg(),
                ]),
              ),
              ...(rnd() < 0.2 ? { doc: arg() } : {}),
              ...(rnd() < 0.3 ? { assign: pick(names) } : {}),
            }
          case 'set':
            return { kind: 'set', variable: pick(names), value: arg() }
          case 'log':
            return { kind: 'log', message: arg() }
          case 'if':
            return {
              kind: 'if',
              branches: Array.from({ length: 1 + Math.floor(rnd() * 2) }, () => ({
                when: pick(exprs),
                steps: steps(depth + 1),
              })),
              ...(rnd() < 0.5 ? { else: steps(depth + 1) } : {}),
            }
          case 'forEach':
            return {
              kind: 'forEach',
              variable: pick(names),
              items: arg(),
              ...(rnd() < 0.4 ? { where: pick(exprs) } : {}),
              steps: steps(depth + 1),
            }
          case 'try':
            return {
              kind: 'try',
              steps: steps(depth + 1),
              catch: steps(depth + 1),
              ...(rnd() < 0.5 ? { errorVariable: pick(names) } : {}),
            }
          case 'transaction':
            return { kind: 'transaction', steps: steps(depth + 1) }
        }
      })
    let actions = 0
    for (let n = 0; n < 300; n++) {
      const w: Workflow = parseWorkflow({
        format: 'dvh-workflow',
        version: 1,
        id: `wf_${n}`,
        name: pick(texts),
        ...(rnd() < 0.3 ? { description: pick(texts) } : {}),
        params:
          rnd() < 0.5
            ? [
                {
                  name: 'p',
                  type: pick(['text', 'list', 'any'] as const),
                  ...(rnd() < 0.7 ? { default: value() } : {}),
                  ...(rnd() < 0.3 ? { description: pick(texts) } : {}),
                },
              ]
            : [],
        steps: steps(0),
        ...(rnd() < 0.2 ? { catalog: 'abc123def456' } : {}),
      })
      const c = canonicalWorkflow(w)
      const printed = printScript(c)
      const back = parseScript(printed)
      if (!back.ok)
        throw new Error(`${back.errors[0]!.message} (line ${back.errors[0]!.line})\n${printed}`)
      expect(back.workflow).toEqual(c)
      expect(JSON.stringify(back.workflow)).toBe(JSON.stringify(c))
      expect(printScript(back.workflow)).toBe(printed)
      actions += JSON.stringify(c).split('"kind":"action"').length - 1
    }
    expect(actions).toBeGreaterThan(100)
  })
})

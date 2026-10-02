/**
 * P9: catalog checks, completions, the sandboxed runner (acceptance: a script
 * reaches nothing but the allowed actions) and the AI script skill.
 */
import { describe, expect, it } from 'vitest'
import {
  ActionRegistry,
  localParticipant,
  z,
  type ActionEffect,
  type ActionGroup,
} from '@genoffice/dvh-actions'
import type { Workflow } from '@genoffice/dvh-workflow'
import {
  checkScript,
  completeAt,
  createDvhScriptSkill,
  lineOfPath,
  printScript,
  runScript,
} from '../src/index'

function fakeDoc() {
  let rows: string[] = []
  const fields: Record<string, unknown> = {}
  const registry = new ActionRegistry()
  const add = (
    name: string,
    group: ActionGroup,
    effect: ActionEffect,
    input: z.ZodType,
    run: (i: never) => unknown,
    objects: (i: never) => number = () => 1,
  ) =>
    registry.register({
      name,
      group,
      summary: `${name} summary`,
      effect,
      input,
      preview: (i) => ({ objects: objects(i as never), files: [], summary: [name], warnings: [] }),
      execute: (i, ctx) => {
        const out = run(i as never)
        if (effect !== 'read') ctx.emit([{ objectId: name, path: '', before: null, after: null }])
        return out
      },
    })
  add(
    'Data.SetField',
    'Data',
    'write',
    z
      .object({
        field: z.string().describe('field name'),
        value: z.union([z.string(), z.number(), z.null()]),
      })
      .strict(),
    (i: { field: string; value: unknown }) => {
      fields[i.field] = i.value
    },
  )
  add(
    'Data.AddRow',
    'Data',
    'write',
    z.object({ text: z.string() }).strict(),
    (i: { text: string }) => {
      if (i.text === 'boom') throw new Error('disk full')
      rows.push(i.text)
      return rows.length
    },
  )
  add(
    'Data.DeleteRows',
    'Data',
    'destructive',
    z.object({ count: z.number().int() }).strict(),
    (i: { count: number }) => {
      rows = rows.slice(i.count)
    },
    (i: { count: number }) => i.count,
  )
  const participant = localParticipant({
    docId: 'doc_a',
    registry,
    checkpoint: () => ({ rows: [...rows], fields: { ...fields } }),
    restore: (cp) => {
      rows = [...(cp as { rows: string[] }).rows]
    },
  })
  return {
    registry,
    participants: new Map([['doc_a', participant]]),
    rows: () => rows,
    fields,
    setRows: (r: string[]) => {
      rows = r
    },
  }
}

const run = (source: string, doc = fakeDoc(), extra = {}) =>
  runScript(source, doc.registry.catalog(), {
    participants: doc.participants,
    docId: 'doc_a',
    ...extra,
  })

describe('checks against the catalog', () => {
  it('reports unknown actions and arguments, missing ones, wrong literal types and risky actions', () => {
    const doc = fakeDoc()
    const { diagnostics } = checkScript(
      [
        'Data.SetField(field:="a", value:=1, colour:="red")',
        'Data.SetField(value:=2)',
        'Data.AddRow(text:=12)',
        'File.ReadAll(path:="C:\\\\Windows")',
        'Data.DeleteRows(count:=3)',
      ].join('\n'),
      doc.registry.catalog(),
    )
    expect(diagnostics.map((d) => [d.line, d.severity, d.message])).toEqual([
      [1, 'error', 'Data.SetField has no argument colour (it takes field, value)'],
      [2, 'error', 'Data.SetField needs field:='],
      [3, 'error', 'Data.AddRow: text expects string'],
      [4, 'error', 'File.ReadAll is not an action of this document'],
      [5, 'warning', 'Data.DeleteRows is destructive: it runs only after the user confirms'],
    ])
  })
})

describe('completions from the catalog', () => {
  const catalog = fakeDoc().registry.catalog()
  const at = (source: string) => completeAt(source.replace('|', ''), source.indexOf('|'), catalog)

  it('action names, then their arguments, then the names in scope', () => {
    const actions = at('Param rows As List\n\nData.Se|')!
    expect(actions.options.map((o) => o.label)).toContain('Data.SetField')
    expect(actions.options.find((o) => o.label === 'Data.SetField')).toMatchObject({
      apply: 'Data.SetField(',
    })
    const args = at('Data.SetField(|')!
    expect(args.options.map((o) => [o.label, o.apply])).toEqual([
      ['field', 'field:='],
      ['value', 'value:='],
    ])
    expect(args.options[0]!.detail).toBe('required string — field name')
    expect(at('Data.SetField(field:="a", v|')!.options.map((o) => o.label)).toEqual(['value'])
    const start = at('Param rows As List\nFor Each row In rows\n    |')!
    expect(start.options.map((o) => o.label)).toEqual(
      expect.arrayContaining(['If … Then', 'Next', 'Data.AddRow', 'row', 'rows']),
    )
    expect(start.options.map((o) => o.label)).not.toContain('Workflow')
    expect(at('|')!.options.map((o) => o.label)).toContain('Param … As …')
    const expr = at('total = 0\nLog UP|')!
    expect(expr.options.map((o) => o.label)).toEqual(
      expect.arrayContaining(['total', 'UPPER', 'INDEX']),
    )
    expect(at('Log "inside a te|')).toBeNull()
  })
})

describe('the sandboxed runner', () => {
  it('runs a script as one transaction and maps failures to lines', async () => {
    const doc = fakeDoc()
    const ok = await run(
      [
        'Param items As List = Json(["a", "b"])',
        '',
        'For Each it In items',
        '    n = Data.AddRow(text:=UPPER(it))',
        'Next',
        'Data.SetField(field:="count", value:=n)',
      ].join('\n'),
      doc,
    )
    expect(ok.ok && ok.result.ok).toBe(true)
    expect(doc.rows()).toEqual(['A', 'B'])
    expect(doc.fields.count).toBe(2)
    if (ok.ok)
      expect(new Set(ok.result.changeSets.map((c) => c.source))).toEqual(new Set(['script']))

    const failed = await run(
      ['Data.AddRow(text:="c")', 'Data.AddRow(text:="boom")'].join('\n'),
      doc,
    )
    expect(failed.ok).toBe(true)
    if (failed.ok) {
      expect(failed.result.ok).toBe(false)
      expect(lineOfPath(failed.lines, failed.result.failedAt)).toBe(2)
    }
    expect(doc.rows()).toEqual(['A', 'B'])
  })

  it('acceptance: a script reaches nothing outside the allowed actions', async () => {
    const doc = fakeDoc()
    const attempts: [string, RegExp][] = [
      ['x = process.env.HOME', /unknown name process\.env\.HOME/],
      ['x = globalThis.fetch', /unknown name globalThis\.fetch/],
      ['x = require("fs")', /unknown function require/],
      ['x = eval("1+1")', /unknown function eval/],
      ['Param r As Record = Json({"a": 1})\nx = r.constructor', /unknown name r\.constructor/],
      ['x = r.__proto__', /unknown name r\.__proto__/],
      ['File.ReadAll(path:="C:\\\\Windows\\\\win.ini")', /not an action of this document/],
      ['Data.DeleteRows(count:=1)', /may not run destructive/],
    ]
    for (const [source, expected] of attempts) {
      const outcome = await run(source, doc)
      const message = outcome.ok
        ? (outcome.result.error ?? 'ran')
        : outcome.diagnostics.map((d) => d.message).join('; ')
      expect(message, source).toMatch(expected)
    }
    // text is never code
    const text = await run('Log "${process.exit(1)}" & "`rm -rf /`"', doc)
    expect(text.ok && text.result.log.at(-1)!.message).toBe('${process.exit(1)}`rm -rf /`')
    // and runs are bounded, even inside Try
    const endless = await run(
      [
        'Param n As List = Json([1,2,3,4,5,6,7,8,9,10])',
        'Try',
        '    For Each a In n',
        '        For Each b In n',
        '            For Each c In n',
        '                For Each d In n',
        '                    x = a',
        '                Next',
        '            Next',
        '        Next',
        '    Next',
        'Catch',
        'End Try',
      ].join('\n'),
      doc,
    )
    expect(endless.ok && endless.result.error).toMatch(/exceeded 10000 steps/)
  })
})

describe('AI writes scripts', () => {
  function skillFor(doc: ReturnType<typeof fakeDoc>, answer: boolean) {
    const asked: string[][] = []
    const saved: Workflow[] = []
    const skill = createDvhScriptSkill({
      registry: () => doc.registry,
      docId: () => 'doc_a',
      participants: () => doc.participants,
      confirm: (lines) => {
        asked.push([...lines])
        return answer
      },
      save: async (w) => {
        saved.push(w)
      },
    })
    const call = (input: Record<string, unknown>) =>
      skill.executeTool({ id: 'c', name: 'propose_script', input })
    return { skill, asked, saved, call }
  }

  it('returns errors with lines, checks, saves and runs scripts', async () => {
    const doc = fakeDoc()
    const { skill, saved, call } = skillFor(doc, true)
    const reference = await skill.executeTool({ id: 'r', name: 'dvh_script_reference', input: {} })
    expect(reference.output).toContain('For Each row In records')
    expect(reference.output).toContain('Data.SetField(field, value) [write]')

    const broken = await call({
      summary: 's',
      mode: 'run',
      source: 'For Each r In rows\nData.AddRow(text:=r)',
    })
    expect(broken.isError).toBe(true)
    expect(broken.output).toContain('line 1: missing Next')
    const wrong = await call({ summary: 's', mode: 'run', source: 'Data.AddRow(txt:="a")' })
    expect(wrong.output).toContain('line 1: Data.AddRow has no argument txt')

    const checked = await call({
      summary: 's',
      mode: 'check',
      source: 'data.addrow(text:="a")'.replace('data.addrow', 'Data.AddRow'),
    })
    expect(JSON.parse(checked.output)).toMatchObject({ ok: true })

    await call({
      summary: 'lưu',
      mode: 'save',
      source: 'Workflow "Thêm"\nId "wf_add"\n\nData.AddRow(text:="x")',
    })
    expect(saved.map((w) => w.id)).toEqual(['wf_add'])
    expect(printScript(saved[0]!)).toContain('Data.AddRow(text:="x")')

    const ran = await call({
      summary: 'thêm hai dòng',
      mode: 'run',
      source: 'Param names As List\n\nFor Each n In names\n    Data.AddRow(text:=n)\nNext',
      params: { names: ['p', 'q'] },
    })
    expect(JSON.parse(ran.output)).toMatchObject({ ok: true, changes: 2 })
    expect(doc.rows()).toEqual(['p', 'q'])

    // a workflow as JSON works the same way
    const json = await call({
      summary: 'json',
      mode: 'run',
      workflow: {
        format: 'dvh-workflow',
        version: 1,
        id: 'wf_j',
        name: 'j',
        params: [],
        steps: [{ kind: 'action', action: 'Data.AddRow', args: { text: { value: 'r' } } }],
      },
    })
    expect(JSON.parse(json.output)).toMatchObject({ ok: true })
    expect(doc.rows()).toEqual(['p', 'q', 'r'])
  })

  it('destructive scripts need the user: declined changes nothing', async () => {
    const doc = fakeDoc()
    doc.setRows(['a', 'b', 'c'])
    const declined = skillFor(doc, false)
    const out = await declined.call({
      summary: 'xoá hai dòng',
      mode: 'run',
      source: 'Data.DeleteRows(count:=2)',
    })
    expect(out.output).toContain('declined')
    expect(declined.asked[0]).toEqual([
      'xoá hai dòng',
      '1. Data.DeleteRows (2 object(s): Data.DeleteRows)',
    ])
    expect(doc.rows()).toEqual(['a', 'b', 'c'])
    const accepted = skillFor(doc, true)
    const done = await accepted.call({
      summary: 'xoá',
      mode: 'run',
      source: 'Data.DeleteRows(count:=2)',
    })
    expect(JSON.parse(done.output)).toMatchObject({ ok: true })
    expect(doc.rows()).toEqual(['c'])
  })
})

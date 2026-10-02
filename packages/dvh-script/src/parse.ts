/**
 * DVH-Script v0 → Workflow Model (P9). A hand-written, line-based parser: one
 * statement per line, blocks closed by `End If` / `Next` / `End Try` /
 * `End Transaction`, keywords in any case.
 *
 *   Workflow "Biên bản"                  ' header lines (all optional)
 *   Id "wf_bb"
 *   Description "…"
 *   Catalog "3f2a…"
 *   Param records As List = Json([]) Description "the rows"
 *
 *   total = 0
 *   For Each row In records Where row.qty > 0
 *       n = Data.AddRow(text:="Hạng mục " & row.name)
 *       total = total + row.qty
 *   Next
 *   If total > 100 Then
 *       Log "big"
 *   ElseIf total > 0 Then
 *       Data.SetField(field:="Total", value:=total) On "doc_x"
 *   Else
 *       Log "nothing"
 *   End If
 *   Try
 *       Transaction
 *           Data.DeleteAll()
 *       End Transaction
 *   Catch err
 *       Log err.message
 *   End Try
 */

import {
  IDENT as WORKFLOW_IDENT,
  workflowSchema,
  type Arg,
  type ParamType,
  type Step,
  type Workflow,
  type WorkflowParam,
} from '@genoffice/dvh-workflow'
import { exprError } from '@genoffice/dvh-template'
import {
  ACTION_NAME,
  IDENT_SRC,
  ScriptError,
  closingParen,
  findKeyword,
  literalOf,
  splitTopLevel,
  stripComment,
} from './lexical'

export interface ParsedScript {
  readonly workflow: Workflow
  /** the source line (1-based) of each step path */
  readonly lines: Readonly<Record<string, number>>
}

export type ParseOutcome =
  | ({ readonly ok: true } & ParsedScript)
  | { readonly ok: false; readonly errors: readonly ScriptError[] }

const TYPES: Record<string, ParamType> = {
  text: 'text',
  number: 'number',
  boolean: 'boolean',
  record: 'record',
  list: 'list',
  any: 'any',
}

type Frame =
  | { kind: 'root'; list: Step[]; base: string }
  | {
      kind: 'if'
      step: Extract<Step, { kind: 'if' }>
      list: Step[]
      base: string
      line: number
      path: string
      elseSeen: boolean
    }
  | { kind: 'forEach'; list: Step[]; base: string; line: number }
  | {
      kind: 'try'
      step: Extract<Step, { kind: 'try' }>
      list: Step[]
      base: string
      line: number
      path: string
      inCatch: boolean
    }
  | { kind: 'transaction'; list: Step[]; base: string; line: number }

const ASSIGN = new RegExp(`^(${IDENT_SRC})\\s*=(?!=)\\s*(.*)$`, 'u')
const CALL = /^([A-Z][A-Za-z]*\.[A-Z][A-Za-z0-9]*)\s*\(/
const FOR_EACH = new RegExp(`^for\\s+each\\s+(${IDENT_SRC})\\s+in\\s+(.+)$`, 'iu')
const CATCH = new RegExp(`^catch(?:\\s+(${IDENT_SRC}))?$`, 'iu')
const NEXT = new RegExp(`^next(?:\\s+(${IDENT_SRC}))?$`, 'iu')
const PARAM = new RegExp(`^param\\s+(${IDENT_SRC})\\s+as\\s+(\\w+)(.*)$`, 'iu')

/** Parses a script; throws the first ScriptError. */
export function compileScript(source: string): ParsedScript {
  const header: {
    name?: string
    id?: string
    description?: string
    catalog?: string
    params: WorkflowParam[]
  } = { params: [] }
  const root: Step[] = []
  const stack: Frame[] = [{ kind: 'root', list: root, base: '' }]
  const lines: Record<string, number> = {}
  let bodyStarted = false

  const rows = source.split(/\r?\n/)
  for (const [index, raw] of rows.entries()) {
    const lineNo = index + 1
    const indent = raw.length - raw.trimStart().length
    const fail = (message: string, at = 0): never => {
      throw new ScriptError(message, lineNo, indent + at + 1)
    }
    let text: string
    try {
      text = stripComment(raw).trim()
    } catch {
      text = raw.trim()
    }
    if (!text) continue
    const top = stack[stack.length - 1]!

    const arg = (piece: string, at: number): Arg => {
      const trimmed = piece.trim()
      if (!trimmed) fail('a value is missing', at)
      let literal: ReturnType<typeof literalOf>
      try {
        literal = literalOf(trimmed)
      } catch (error) {
        fail(error instanceof Error ? error.message : String(error), at)
      }
      if (literal) return literal
      const problem = exprError(trimmed)
      if (problem) fail(`bad expression: ${problem}`, at)
      return { expr: trimmed }
    }
    const expr = (piece: string, at: number): string => {
      const trimmed = piece.trim()
      if (!trimmed) fail('an expression is missing', at)
      const problem = exprError(trimmed)
      if (problem) fail(`bad expression: ${problem}`, at)
      return trimmed
    }
    const push = (step: Step) => {
      const path = top.base ? `${top.base}.${top.list.length}` : String(top.list.length)
      top.list.push(step)
      lines[path] = lineNo
      return path
    }
    const stringLiteral = (piece: string, at: number): string => {
      const value = literalOf(piece)?.value
      if (typeof value !== 'string') fail('expected text in double quotes', at)
      return value as string
    }

    // ----- header -----
    // a header word stands alone (Workflow.Run(...) is a statement)
    const word = /^(\w+)(?=\s)/.exec(text)?.[1]?.toLowerCase()
    if (!bodyStarted && !ASSIGN.test(text)) {
      const rest = text.slice(word?.length ?? 0).trim()
      const at = text.length - rest.length
      if (word === 'workflow' || word === 'id' || word === 'description' || word === 'catalog') {
        const key = word === 'workflow' ? 'name' : word
        if (header[key] !== undefined) fail(`${word} is given twice`)
        header[key] = stringLiteral(rest, at)
        continue
      }
      if (word === 'param') {
        const m = PARAM.exec(text)
        if (!m) fail('expected: Param <name> As <Text|Number|Boolean|Record|List|Any> [= value]')
        const type = TYPES[m![2]!.toLowerCase()]
        if (!type) fail(`unknown type ${m![2]}`, text.indexOf(m![2]!))
        if (header.params.some((p) => p.name === m![1]))
          fail(`parameter ${m![1]} is declared twice`)
        let tail = m![3]!.trim()
        const tailAt = text.length - tail.length
        const param: WorkflowParam = { name: m![1]!, type: type! }
        if (tail.startsWith('=')) {
          tail = tail.slice(1).trim()
          const cut = findKeyword(` ${tail} `, 'Description')
          const valueText = cut < 0 ? tail : tail.slice(0, cut - 1)
          const value = literalOf(valueText)
          if (!value)
            fail(
              'a parameter default must be a literal (text, number, True/False/Null or Json(...))',
              tailAt,
            )
          param.default = value!.value
          tail = cut < 0 ? '' : tail.slice(cut - 1).trim()
        }
        if (tail) {
          const d = /^description\s+(.*)$/i.exec(tail)
          if (!d) fail(`unexpected "${tail}"`, tailAt)
          param.description = stringLiteral(d![1]!, tailAt)
        }
        header.params.push(param)
        continue
      }
    }
    bodyStarted = true

    // ----- simple statements -----
    const call = (callText: string, at: number, assign?: string): Step => {
      const m = CALL.exec(callText)!
      const open = m[0].length - 1
      const close = closingParen(callText, open)
      if (close < 0) fail('missing ")"', at + open)
      const inside = callText.slice(open + 1, close)
      const args: Record<string, Arg> = {}
      if (inside.trim()) {
        for (const part of splitTopLevel(inside)) {
          const partAt = at + open + 1 + part.offset
          const named = new RegExp(`^\\s*(${IDENT_SRC})\\s*:=([\\s\\S]*)$`, 'u').exec(part.text)
          if (!named) fail('arguments are named: name:=value', partAt)
          if (args[named![1]!]) fail(`argument ${named![1]} is given twice`, partAt)
          args[named![1]!] = arg(named![2]!, partAt + part.text.indexOf(':=') + 2)
        }
      }
      const step: Extract<Step, { kind: 'action' }> = { kind: 'action', action: m[1]!, args }
      const tail = callText.slice(close + 1).trim()
      if (tail) {
        const on = /^on\s+(.+)$/i.exec(tail)
        if (!on) fail(`unexpected "${tail}" after the call`, at + close + 1)
        step.doc = arg(on![1]!, at + callText.length - on![1]!.length)
      }
      if (assign) step.assign = assign
      if (!ACTION_NAME.test(step.action)) fail(`${step.action} is not an action name`, at)
      return step
    }

    // an assignment first: a variable may be called Else or Next
    const assignment = ASSIGN.exec(text)
    if (assignment) {
      const name = assignment[1]!
      if (!WORKFLOW_IDENT.test(name)) fail(`${name} is not a valid name`)
      const rhs = assignment[2]!
      const rhsAt = text.length - rhs.length
      if (CALL.test(rhs)) {
        push(call(rhs, rhsAt, name))
      } else {
        push({ kind: 'set', variable: name, value: arg(rhs, rhsAt) })
      }
      continue
    }
    // ----- block ends and middles -----
    const lower = text.toLowerCase().replace(/\s+/g, ' ')
    if (lower === 'end if') {
      if (top.kind !== 'if') fail('End If without If')
      stack.pop()
      continue
    }
    if (lower === 'end try') {
      if (top.kind !== 'try') fail('End Try without Try')
      if (!(top as { inCatch: boolean }).inCatch) fail('Try needs a Catch')
      stack.pop()
      continue
    }
    if (lower === 'end transaction') {
      if (top.kind !== 'transaction') fail('End Transaction without Transaction')
      stack.pop()
      continue
    }
    const next = NEXT.exec(text)
    if (next) {
      if (top.kind !== 'forEach') fail('Next without For Each')
      stack.pop()
      continue
    }
    if (/^else\s*if\s/i.test(text) || /^elseif\s/i.test(text)) {
      if (top.kind !== 'if' || top.elseSeen) fail('ElseIf outside If')
      const frame = top as Extract<Frame, { kind: 'if' }>
      const m = /^else\s*if\s+(.*)\s+then$/i.exec(text)
      if (!m) fail('expected: ElseIf <condition> Then')
      const branch = { when: expr(m![1]!, text.indexOf(m![1]!)), steps: [] as Step[] }
      frame.step.branches.push(branch)
      frame.list = branch.steps
      frame.base = `${frame.path}.then${frame.step.branches.length - 1}`
      continue
    }
    if (lower === 'else') {
      if (top.kind !== 'if' || top.elseSeen) fail('Else outside If')
      const frame = top as Extract<Frame, { kind: 'if' }>
      frame.step.else = []
      frame.list = frame.step.else
      frame.base = `${frame.path}.else`
      frame.elseSeen = true
      continue
    }
    const catchMatch = CATCH.exec(text)
    if (catchMatch) {
      if (top.kind !== 'try' || top.inCatch) fail('Catch outside Try')
      const frame = top as Extract<Frame, { kind: 'try' }>
      if (catchMatch[1]) frame.step.errorVariable = catchMatch[1]
      frame.list = frame.step.catch
      frame.base = `${frame.path}.catch`
      frame.inCatch = true
      continue
    }

    // ----- block starts -----
    if (/^if\s/i.test(text)) {
      const m = /^if\s+(.*)\s+then$/i.exec(text)
      if (!m) fail('expected: If <condition> Then')
      const step: Extract<Step, { kind: 'if' }> = {
        kind: 'if',
        branches: [{ when: expr(m![1]!, 3), steps: [] }],
      }
      const path = push(step)
      stack.push({
        kind: 'if',
        step,
        list: step.branches[0]!.steps,
        base: `${path}.then0`,
        line: lineNo,
        path,
        elseSeen: false,
      })
      continue
    }
    const forEach = FOR_EACH.exec(text)
    if (forEach) {
      const tail = forEach[2]!
      const tailAt = text.length - tail.length
      const cut = findKeyword(` ${tail} `, 'Where')
      const itemsText = cut < 0 ? tail : tail.slice(0, cut - 1)
      const step: Extract<Step, { kind: 'forEach' }> = {
        kind: 'forEach',
        variable: forEach[1]!,
        items: arg(itemsText, tailAt),
        ...(cut >= 0
          ? {
              where: expr(
                tail
                  .slice(cut - 1)
                  .trim()
                  .slice(5),
                tailAt + cut + 5,
              ),
            }
          : {}),
        steps: [],
      }
      const path = push(step)
      stack.push({ kind: 'forEach', list: step.steps, base: `${path}.body`, line: lineNo })
      continue
    }
    if (lower === 'try') {
      const step: Extract<Step, { kind: 'try' }> = { kind: 'try', steps: [], catch: [] }
      const path = push(step)
      stack.push({
        kind: 'try',
        step,
        list: step.steps,
        base: `${path}.body`,
        line: lineNo,
        path,
        inCatch: false,
      })
      continue
    }
    if (lower === 'transaction') {
      const step: Extract<Step, { kind: 'transaction' }> = { kind: 'transaction', steps: [] }
      const path = push(step)
      stack.push({ kind: 'transaction', list: step.steps, base: `${path}.body`, line: lineNo })
      continue
    }

    if (CALL.test(text)) {
      push(call(text, 0))
      continue
    }
    if (/^log(\s|$)/i.test(text)) {
      const rest = text.slice(3)
      push({ kind: 'log', message: arg(rest, 3) })
      continue
    }
    fail(`cannot read "${text.length > 40 ? `${text.slice(0, 40)}…` : text}"`)
  }

  const open = stack[stack.length - 1]!
  if (open.kind !== 'root') {
    const closing = {
      if: 'End If',
      forEach: 'Next',
      try: 'End Try',
      transaction: 'End Transaction',
    }[open.kind]
    throw new ScriptError(`missing ${closing}`, open.line, 1)
  }
  const workflow = {
    format: 'dvh-workflow' as const,
    version: 1 as const,
    id: header.id ?? 'wf_script',
    name: header.name ?? 'Script',
    ...(header.description !== undefined ? { description: header.description } : {}),
    params: header.params,
    steps: root,
    ...(header.catalog !== undefined ? { catalog: header.catalog } : {}),
  }
  const checked = workflowSchema.safeParse(workflow)
  if (!checked.success) {
    const issue = checked.error.issues[0]!
    const path = stepPathOf(issue.path)
    throw new ScriptError(issue.message, (path && lines[path]) || 1, 1)
  }
  return { workflow: checked.data, lines }
}

/** The step path (`2.body.0`) of a zod issue path into a workflow, if it is inside the steps. */
function stepPathOf(issuePath: readonly PropertyKey[]): string | null {
  if (issuePath[0] !== 'steps') return null
  const out: string[] = []
  let i = 1
  while (typeof issuePath[i] === 'number') {
    out.push(String(issuePath[i]))
    const key = issuePath[i + 1]
    if (
      (key === 'steps' || key === 'catch' || key === 'else') &&
      typeof issuePath[i + 2] === 'number'
    ) {
      out.push(key === 'steps' ? 'body' : key)
      i += 2
      continue
    }
    if (
      key === 'branches' &&
      typeof issuePath[i + 2] === 'number' &&
      issuePath[i + 3] === 'steps' &&
      typeof issuePath[i + 4] === 'number'
    ) {
      out.push(`then${String(issuePath[i + 2])}`)
      i += 4
      continue
    }
    break
  }
  return out.join('.')
}

/** Parses a script and reports every problem as an error list instead of throwing. */
export function parseScript(source: string): ParseOutcome {
  try {
    return { ok: true, ...compileScript(source) }
  } catch (error) {
    if (error instanceof ScriptError) return { ok: false, errors: [error] }
    return {
      ok: false,
      errors: [new ScriptError(error instanceof Error ? error.message : String(error), 1, 1)],
    }
  }
}

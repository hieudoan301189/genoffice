/**
 * Completions for the script editor (P9), from the action catalog: action
 * names, the arguments of the action being called, keywords at the start of a
 * statement, and the names in scope (parameters, variables, loop variables)
 * and expression functions inside expressions. Editor-agnostic: the Docs
 * editor maps them to CodeMirror completions.
 */

import { argumentsOf, type ScriptCatalogEntry } from './check'
import { IDENT_SRC, closingParen, stripComment } from './lexical'

export interface ScriptCompletion {
  readonly label: string
  readonly kind: 'action' | 'argument' | 'keyword' | 'variable' | 'function'
  readonly detail?: string
  /** the text to insert when it differs from the label */
  readonly apply?: string
}

export interface CompletionResult {
  /** offset in the source where the completed word starts */
  readonly from: number
  readonly options: readonly ScriptCompletion[]
}

const STATEMENT_KEYWORDS = [
  'If … Then',
  'ElseIf … Then',
  'Else',
  'End If',
  'For Each … In …',
  'Next',
  'Try',
  'Catch',
  'End Try',
  'Transaction',
  'End Transaction',
  'Log',
]
const HEADER_KEYWORDS = ['Workflow', 'Id', 'Description', 'Catalog', 'Param … As …']
const FUNCTIONS = [
  'IF',
  'AND',
  'OR',
  'NOT',
  'LEN',
  'UPPER',
  'LOWER',
  'TRIM',
  'LEFT',
  'RIGHT',
  'MID',
  'CONTAINS',
  'ISBLANK',
  'ROUND',
  'ABS',
  'MIN',
  'MAX',
  'SUM',
  'CONCAT',
  'TEXT',
  'PAD',
]

const keywordApply = (k: string) => k.replace(/ …/g, ' ').replace(/ $/, ' ')

/** The names a script binds: parameters, assignments, loop and error variables. */
export function namesInScript(source: string): string[] {
  const names = new Set<string>(['INDEX'])
  const ident = new RegExp(`^\\s*(${IDENT_SRC})\\s*=(?!=)`, 'u')
  const param = new RegExp(`^\\s*param\\s+(${IDENT_SRC})`, 'iu')
  const loop = new RegExp(`^\\s*for\\s+each\\s+(${IDENT_SRC})`, 'iu')
  const caught = new RegExp(`^\\s*catch\\s+(${IDENT_SRC})`, 'iu')
  for (const line of source.split(/\r?\n/)) {
    for (const re of [param, loop, caught, ident]) {
      const m = re.exec(line)
      if (m) {
        names.add(m[1]!)
        break
      }
    }
  }
  return [...names]
}

/** Completions at `offset` (the caret) of `source`. */
export function completeAt(
  source: string,
  offset: number,
  catalog: readonly ScriptCatalogEntry[],
): CompletionResult | null {
  const lineStart = source.lastIndexOf('\n', offset - 1) + 1
  const before = source.slice(lineStart, offset)
  // nothing to complete inside a comment or a text
  if (stripComment(before).length < before.trimEnd().length) return null
  if ((before.match(/"/g)?.length ?? 0) % 2 === 1) return null

  const word = /[\p{L}\p{N}_.]*$/u.exec(before)![0]
  const from = offset - word.length
  const head = before.slice(0, before.length - word.length)

  // inside the parentheses of an action call: its arguments
  const call = /([A-Z][A-Za-z]*\.[A-Z][A-Za-z0-9]*)\s*\(/g
  let open: { name: string; at: number } | null = null
  for (const m of before.matchAll(call)) {
    const at = m.index! + m[0].length - 1
    if (closingParen(before, at) < 0) open = { name: m[1]!, at }
  }
  if (open && /[(,]\s*$/.test(head) && !word.includes('.')) {
    const entry = catalog.find((e) => e.name === open!.name)
    if (!entry) return null
    const used = new Set(
      [...before.slice(open.at).matchAll(new RegExp(`(${IDENT_SRC})\\s*:=`, 'gu'))].map(
        (m) => m[1]!,
      ),
    )
    const { names, required, types, descriptions } = argumentsOf(entry)
    return {
      from,
      options: names
        .filter((n) => !used.has(n))
        .map((n) => ({
          label: n,
          kind: 'argument' as const,
          detail: `${required.includes(n) ? 'required' : 'optional'}${types[n] ? ` ${types[n]!.join('|')}` : ''}${descriptions[n] ? ` — ${descriptions[n]}` : ''}`,
          apply: `${n}:=`,
        })),
    }
  }

  const atStatementStart = /^\s*$/.test(head)
  const afterAssign = new RegExp(`^\\s*${IDENT_SRC}\\s*=\\s*$`, 'u').test(head)
  const actions: ScriptCompletion[] = catalog.map((e) => ({
    label: e.name,
    kind: 'action',
    detail: `${e.effect}${e.summary ? ` — ${e.summary}` : ''}`,
    apply: `${e.name}(`,
  }))
  // an action name being typed (Group. or Group.Ve)
  if (/^[A-Z][A-Za-z]*\.[A-Za-z0-9]*$/.test(word) && (atStatementStart || afterAssign))
    return { from, options: actions }

  const names: ScriptCompletion[] = namesInScript(source).map((n) => ({
    label: n,
    kind: 'variable',
  }))
  if (atStatementStart) {
    // header keywords only before the first statement
    const inHeader = source
      .slice(0, lineStart)
      .split(/\r?\n/)
      .map((line) => stripComment(line).trim())
      .every((line) => !line || /^(workflow|id|description|catalog|param)\s/i.test(line))
    return {
      from,
      options: [
        ...(inHeader ? HEADER_KEYWORDS : []).map((k) => ({
          label: k,
          kind: 'keyword' as const,
          apply: keywordApply(k),
        })),
        ...STATEMENT_KEYWORDS.map((k) => ({
          label: k,
          kind: 'keyword' as const,
          apply: keywordApply(k),
        })),
        ...actions,
        ...names,
      ],
    }
  }
  // inside an expression: names and functions (and actions right after "x =")
  return {
    from,
    options: [
      ...(afterAssign ? actions : []),
      ...names,
      ...FUNCTIONS.map((f) => ({ label: f, kind: 'function' as const, apply: `${f}(` })),
    ],
  }
}

/**
 * Safe expressions for Smart Templates (P6): conditions, filters and file-name
 * rules. Extends the arithmetic of Sheets' dvh-evaluate.ts with values,
 * comparisons, logic, text and a small function library — parsed into a tree
 * and interpreted, never `eval`ed, with a step budget.
 *
 * Grammar (Excel-like):
 *   expr    := or
 *   or      := and (OR and)*            and := not (AND not)*
 *   not     := NOT not | compare        compare := concat ((= <> != < <= > >=) concat)?
 *   concat  := add (& add)*             add := mul ((+|-) mul)*
 *   mul     := unary ((*|/|%) unary)*   unary := (+|-) unary | power
 *   power   := primary (^ unary)?
 *   primary := number | "text" | 'text' | TRUE | FALSE | NULL | name | [Title]
 *            | name(args) | ( expr )
 * Names are dotted (`Project.Name`, `row.qty`); `[Column title]` reads the
 * current record by title.
 */

import type { Scalar } from '@genoffice/dvh-model'

export type ExprValue = Scalar

export interface ExprScope {
  /** a dotted name (`Project.Name`, `row.qty`, `qty`); undefined when unknown */
  lookup(name: string): ExprValue | undefined
  /** `[Title]`: a column of the current record by title */
  column?(title: string): ExprValue | undefined
}

export class ExprError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'ExprError'
  }
}

export type ExprNode =
  | { k: 'lit'; v: ExprValue }
  | { k: 'name'; name: string }
  | { k: 'col'; title: string }
  | { k: 'un'; op: string; a: ExprNode }
  | { k: 'bin'; op: string; a: ExprNode; b: ExprNode }
  | { k: 'call'; fn: string; args: ExprNode[] }

const MAX_LENGTH = 2000
const MAX_STEPS = 10_000

interface Token {
  t: 'num' | 'str' | 'name' | 'col' | 'op' | 'end'
  v: string
}

function tokenize(source: string): Token[] {
  if (source.length > MAX_LENGTH) throw new ExprError('expression too long')
  const out: Token[] = []
  let i = 0
  while (i < source.length) {
    const ch = source[i]!
    if (/\s/.test(ch)) {
      i++
      continue
    }
    if (/[0-9.]/.test(ch)) {
      const m = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(source.slice(i))
      if (!m) throw new ExprError(`bad number at ${i}`)
      out.push({ t: 'num', v: m[0] })
      i += m[0].length
      continue
    }
    if (ch === '"' || ch === "'") {
      let j = i + 1
      let text = ''
      for (;;) {
        if (j >= source.length) throw new ExprError('unterminated text')
        if (source[j] === ch) {
          // a doubled quote is a literal quote, as in Excel
          if (source[j + 1] === ch) {
            text += ch
            j += 2
            continue
          }
          break
        }
        text += source[j]
        j++
      }
      out.push({ t: 'str', v: text })
      i = j + 1
      continue
    }
    if (ch === '[') {
      const end = source.indexOf(']', i)
      if (end < 0) throw new ExprError('unterminated [column]')
      out.push({ t: 'col', v: source.slice(i + 1, end).trim() })
      i = end + 1
      continue
    }
    const name = /^[\p{L}_][\p{L}\p{N}_]*(?:\.[\p{L}_][\p{L}\p{N}_]*)*/u.exec(source.slice(i))
    if (name) {
      out.push({ t: 'name', v: name[0] })
      i += name[0].length
      continue
    }
    const op = /^(?:<>|!=|<=|>=|[-+*/%^&=<>(),])/.exec(source.slice(i))
    if (op) {
      out.push({ t: 'op', v: op[0] })
      i += op[0].length
      continue
    }
    throw new ExprError(`unexpected "${ch}" at ${i}`)
  }
  out.push({ t: 'end', v: '' })
  return out
}

/** Parses an expression (a leading `=` is allowed) into a tree; throws ExprError. */
export function parseExpr(input: string): ExprNode {
  const tokens = tokenize(input.trim().replace(/^=/, ''))
  let pos = 0
  const peek = () => tokens[pos]!
  const isOp = (v: string) => peek().t === 'op' && peek().v === v
  const isWord = (v: string) => peek().t === 'name' && peek().v.toUpperCase() === v
  const expect = (v: string) => {
    if (!isOp(v)) throw new ExprError(`expected "${v}"`)
    pos++
  }
  const or = (): ExprNode => {
    let a = and()
    while (isWord('OR')) {
      pos++
      a = { k: 'bin', op: 'or', a, b: and() }
    }
    return a
  }
  const and = (): ExprNode => {
    let a = not()
    while (isWord('AND')) {
      pos++
      a = { k: 'bin', op: 'and', a, b: not() }
    }
    return a
  }
  const not = (): ExprNode => {
    if (isWord('NOT') && tokens[pos + 1]?.v !== '(') {
      pos++
      return { k: 'un', op: 'not', a: not() }
    }
    return compare()
  }
  const compare = (): ExprNode => {
    const a = concat()
    const t = peek()
    if (t.t === 'op' && ['=', '<>', '!=', '<', '<=', '>', '>='].includes(t.v)) {
      pos++
      return { k: 'bin', op: t.v === '!=' ? '<>' : t.v, a, b: concat() }
    }
    return a
  }
  const concat = (): ExprNode => {
    let a = add()
    while (isOp('&')) {
      pos++
      a = { k: 'bin', op: '&', a, b: add() }
    }
    return a
  }
  const add = (): ExprNode => {
    let a = mul()
    while (isOp('+') || isOp('-')) {
      const op = tokens[pos++]!.v
      a = { k: 'bin', op, a, b: mul() }
    }
    return a
  }
  const mul = (): ExprNode => {
    let a = unary()
    while (isOp('*') || isOp('/') || isOp('%')) {
      const op = tokens[pos++]!.v
      a = { k: 'bin', op, a, b: unary() }
    }
    return a
  }
  const unary = (): ExprNode => {
    if (isOp('+') || isOp('-')) {
      const op = tokens[pos++]!.v
      return { k: 'un', op, a: unary() }
    }
    return power()
  }
  const power = (): ExprNode => {
    const a = primary()
    if (isOp('^')) {
      pos++
      return { k: 'bin', op: '^', a, b: unary() }
    }
    return a
  }
  const primary = (): ExprNode => {
    const t = peek()
    if (t.t === 'num') {
      pos++
      return { k: 'lit', v: Number(t.v) }
    }
    if (t.t === 'str') {
      pos++
      return { k: 'lit', v: t.v }
    }
    if (t.t === 'col') {
      pos++
      return { k: 'col', title: t.v }
    }
    if (isOp('(')) {
      pos++
      const inner = or()
      expect(')')
      return inner
    }
    if (t.t === 'name') {
      pos++
      const upper = t.v.toUpperCase()
      if (isOp('(')) {
        pos++
        const args: ExprNode[] = []
        if (!isOp(')')) {
          args.push(or())
          while (isOp(',')) {
            pos++
            args.push(or())
          }
        }
        expect(')')
        if (!(upper in FUNCTIONS)) throw new ExprError(`unknown function ${t.v}`)
        return { k: 'call', fn: upper, args }
      }
      if (upper === 'TRUE') return { k: 'lit', v: true }
      if (upper === 'FALSE') return { k: 'lit', v: false }
      if (upper === 'NULL') return { k: 'lit', v: null }
      return { k: 'name', name: t.v }
    }
    throw new ExprError(t.t === 'end' ? 'unexpected end' : `unexpected "${t.v}"`)
  }
  const tree = or()
  if (peek().t !== 'end') throw new ExprError(`unexpected "${peek().v}"`)
  return tree
}

const num = (v: ExprValue): number => {
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  if (v === null || v === '') return 0
  const n = Number(String(v).replace(/\s/g, ''))
  if (!Number.isFinite(n)) throw new ExprError(`"${v}" is not a number`)
  return n
}
const text = (v: ExprValue): string =>
  v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)
export const truthy = (v: ExprValue): boolean =>
  v !== null && v !== false && v !== 0 && v !== '' && v !== 'FALSE'

function compareValues(a: ExprValue, b: ExprValue): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (
    (typeof a === 'number' || typeof b === 'number') &&
    a !== null &&
    b !== null &&
    a !== '' &&
    b !== ''
  ) {
    const na = Number(a)
    const nb = Number(b)
    if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb
  }
  return text(a).localeCompare(text(b), undefined, { sensitivity: 'accent' })
}

const FUNCTIONS: Record<string, (args: ExprValue[]) => ExprValue> = {
  IF: ([c, a = null, b = null]) => (truthy(c ?? null) ? a : b),
  AND: (args) => args.every((a) => truthy(a)),
  OR: (args) => args.some((a) => truthy(a)),
  NOT: ([a = null]) => !truthy(a),
  LEN: ([a = null]) => text(a).length,
  UPPER: ([a = null]) => text(a).toUpperCase(),
  LOWER: ([a = null]) => text(a).toLowerCase(),
  TRIM: ([a = null]) => text(a).trim().replace(/\s+/g, ' '),
  LEFT: ([a = null, n = 1]) => text(a).slice(0, Math.max(0, num(n))),
  RIGHT: ([a = null, n = 1]) => (num(n) <= 0 ? '' : text(a).slice(-num(n))),
  MID: ([a = null, s = 1, n = 1]) => text(a).substr(Math.max(0, num(s) - 1), Math.max(0, num(n))),
  CONTAINS: ([a = null, b = null]) => text(a).toLowerCase().includes(text(b).toLowerCase()),
  ISBLANK: ([a = null]) => a === null || a === '',
  ROUND: ([a = null, d = 0]) => {
    const f = 10 ** Math.trunc(num(d))
    return Math.round(num(a) * f) / f
  },
  ABS: ([a = null]) => Math.abs(num(a)),
  MIN: (args) => Math.min(...args.map(num)),
  MAX: (args) => Math.max(...args.map(num)),
  SUM: (args) => args.reduce<number>((s, a) => s + num(a), 0),
  CONCAT: (args) => args.map(text).join(''),
  TEXT: ([a = null, digits = null]) =>
    digits === null ? text(a) : num(a).toFixed(Math.max(0, Math.min(15, num(digits)))),
  PAD: ([a = null, width = 2, fill = '0']) => text(a).padStart(num(width), text(fill) || '0'),
}

/** Evaluates a parsed expression against a scope; unknown names are errors, not blanks. */
export function evaluateTree(tree: ExprNode, scope: ExprScope): ExprValue {
  let steps = 0
  const ev = (n: ExprNode): ExprValue => {
    if (++steps > MAX_STEPS) throw new ExprError('expression too complex')
    switch (n.k) {
      case 'lit':
        return n.v
      case 'name': {
        const v = scope.lookup(n.name)
        if (v === undefined) throw new ExprError(`unknown name ${n.name}`)
        return v
      }
      case 'col': {
        const v = scope.column?.(n.title)
        if (v === undefined) throw new ExprError(`unknown column [${n.title}]`)
        return v
      }
      case 'un': {
        const a = ev(n.a)
        if (n.op === 'not') return !truthy(a)
        return n.op === '-' ? -num(a) : num(a)
      }
      case 'bin': {
        if (n.op === 'and') return truthy(ev(n.a)) && truthy(ev(n.b))
        if (n.op === 'or') return truthy(ev(n.a)) || truthy(ev(n.b))
        const a = ev(n.a)
        const b = ev(n.b)
        switch (n.op) {
          case '&':
            return text(a) + text(b)
          case '+':
            return num(a) + num(b)
          case '-':
            return num(a) - num(b)
          case '*':
            return num(a) * num(b)
          case '/': {
            if (num(b) === 0) throw new ExprError('division by zero')
            return num(a) / num(b)
          }
          case '%':
            return num(a) % num(b)
          case '^':
            return num(a) ** num(b)
          case '=':
            return compareValues(a, b) === 0
          case '<>':
            return compareValues(a, b) !== 0
          case '<':
            return compareValues(a, b) < 0
          case '<=':
            return compareValues(a, b) <= 0
          case '>':
            return compareValues(a, b) > 0
          case '>=':
            return compareValues(a, b) >= 0
        }
        throw new ExprError(`unknown operator ${n.op}`)
      }
      case 'call':
        // IF evaluates lazily so the unused branch cannot fail the whole expression
        if (n.fn === 'IF') {
          return truthy(ev(n.args[0] ?? { k: 'lit', v: null }))
            ? n.args[1]
              ? ev(n.args[1])
              : null
            : n.args[2]
              ? ev(n.args[2])
              : null
        }
        return FUNCTIONS[n.fn]!(n.args.map(ev))
    }
  }
  const result = ev(tree)
  if (typeof result === 'number' && !Number.isFinite(result))
    throw new ExprError('not a finite number')
  return result
}

/** Parses and evaluates in one go. */
export function evaluateExpr(source: string, scope: ExprScope): ExprValue {
  return evaluateTree(parseExpr(source), scope)
}

/** Checks an expression's syntax; returns the error message or null. */
export function exprError(source: string): string | null {
  try {
    parseExpr(source)
    return null
  } catch (error) {
    return error instanceof Error ? error.message : String(error)
  }
}

/**
 * A text with `{expression}` holes, e.g. a file-name rule
 * `BB-{row.code}-{PAD(INDEX, 3)}`; `{{` and `}}` are literal braces.
 */
export function renderTemplateText(rule: string, scope: ExprScope): string {
  let out = ''
  let i = 0
  while (i < rule.length) {
    const ch = rule[i]!
    if (ch === '{' && rule[i + 1] === '{') {
      out += '{'
      i += 2
    } else if (ch === '}' && rule[i + 1] === '}') {
      out += '}'
      i += 2
    } else if (ch === '{') {
      const end = rule.indexOf('}', i)
      if (end < 0) throw new ExprError('unterminated {expression}')
      out += text(evaluateExpr(rule.slice(i + 1, end), scope))
      i = end + 1
    } else {
      out += ch
      i++
    }
  }
  return out
}

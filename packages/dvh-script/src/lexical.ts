/**
 * The lexical layer of DVH-Script shared by the parser and the printer:
 * comments, top-level scanning outside strings and brackets, and literals.
 *
 * - `'` starts a comment (outside a double-quoted string), as in VBA, so text
 *   in scripts is always double-quoted ("a ""quoted"" word").
 * - A literal argument is a string, a number, True/False/Null, or
 *   `Json(<json>)` for lists, records and text with line breaks. The printer
 *   writes JSON strings without `"` or `'` inside (", '), so the
 *   scanner never needs to know JSON's own escapes.
 */

import type { Json } from '@genoffice/dvh-workflow'

export class ScriptError extends Error {
  constructor(
    message: string,
    /** 1-based */
    readonly line: number,
    /** 1-based */
    readonly column: number,
  ) {
    super(message)
    this.name = 'ScriptError'
  }
}

export const IDENT_SRC = '[\\p{L}_][\\p{L}\\p{N}_]*'
export const IDENT = new RegExp(`^${IDENT_SRC}$`, 'u')
export const ACTION_NAME = /^[A-Z][A-Za-z]*\.[A-Z][A-Za-z0-9]*$/

/**
 * Walks a text, calling `visit(i, depth)` for every character outside a
 * string literal; returns false from visit to stop. Throws on an unterminated
 * string (with its offset).
 */
export function scanTopLevel(
  text: string,
  visit: (index: number, depth: number) => boolean | void,
): void {
  let depth = 0
  let i = 0
  while (i < text.length) {
    const ch = text[i]!
    if (ch === '"') {
      let j = i + 1
      for (;;) {
        if (j >= text.length) throw new RangeError(String(i))
        if (text[j] === '"') {
          if (text[j + 1] === '"') {
            j += 2
            continue
          }
          break
        }
        j++
      }
      i = j + 1
      continue
    }
    if (ch === '(' || ch === '[' || ch === '{') depth++
    if (visit(i, depth) === false) return
    if (ch === ')' || ch === ']' || ch === '}') depth--
    i++
  }
}

/** The line without its comment (a `'` outside strings), right-trimmed. */
export function stripComment(line: string): string {
  let cut = line.length
  try {
    // a ' inside [Column title] belongs to the title
    scanTopLevel(line, (i, depth) => {
      if (line[i] === "'" && depth === 0) {
        cut = i
        return false
      }
      return true
    })
  } catch {
    // an unterminated string: the parser reports it where it is used
  }
  return line.slice(0, cut).trimEnd()
}

/** Splits at top-level commas (outside strings and brackets). */
export function splitTopLevel(text: string, separator = ','): { text: string; offset: number }[] {
  const parts: { text: string; offset: number }[] = []
  let start = 0
  scanTopLevel(text, (i, depth) => {
    if (depth === 0 && text[i] === separator) {
      parts.push({ text: text.slice(start, i), offset: start })
      start = i + 1
    }
  })
  parts.push({ text: text.slice(start), offset: start })
  return parts
}

/** Index of the `)` closing the `(` at `open`, or -1. */
export function closingParen(text: string, open: number): number {
  let found = -1
  scanTopLevel(text.slice(open), (i, depth) => {
    if (text[open + i] === ')' && depth === 1) {
      found = open + i
      return false
    }
    return true
  })
  return found
}

/**
 * Finds a keyword (`Where`, `On`) at the top level, surrounded by spaces;
 * returns its offset or -1.
 */
export function findKeyword(text: string, keyword: string): number {
  const re = new RegExp(`\\s${keyword}\\s`, 'iy')
  let found = -1
  scanTopLevel(text, (i, depth) => {
    if (depth !== 0) return true
    re.lastIndex = i
    if (re.test(text)) {
      found = i
      return false
    }
    return true
  })
  return found
}

// ---------- literals ----------

const NUMBER = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/
// eslint-disable-next-line no-control-regex -- line breaks and control characters force the Json() form
const NEEDS_JSON = /[\u0000-\u001f\u007f\u2028\u2029]/

/** A JSON text with no `"` or `'` inside its strings (see the module note). */
export function scriptJson(value: Json): string {
  if (value === null || typeof value === 'boolean') return String(value)
  if (typeof value === 'number') return JSON.stringify(value)
  if (typeof value === 'string') {
    let out = '"'
    for (const ch of value) {
      const code = ch.codePointAt(0)!
      if (ch === '"') out += '\\u0022'
      else if (ch === "'") out += '\\u0027'
      else if (ch === '\\') out += '\\\\'
      else if (code < 0x20 || code === 0x7f || code === 0x2028 || code === 0x2029)
        out += `\\u${code.toString(16).padStart(4, '0')}`
      else out += ch
    }
    return `${out}"`
  }
  if (Array.isArray(value)) return `[${value.map(scriptJson).join(', ')}]`
  return `{${Object.entries(value)
    .map(([k, v]) => `${scriptJson(k)}: ${scriptJson(v)}`)
    .join(', ')}}`
}

/** How a literal value is written in a script. */
export function printLiteral(value: Json): string {
  if (value === null) return 'Null'
  if (value === true) return 'True'
  if (value === false) return 'False'
  if (typeof value === 'number') {
    const text = JSON.stringify(value)
    return NUMBER.test(text) ? text : `Json(${text})`
  }
  if (typeof value === 'string' && !NEEDS_JSON.test(value)) return `"${value.replace(/"/g, '""')}"`
  return `Json(${scriptJson(value)})`
}

/**
 * The literal a whole text spells, or undefined when it is not one (then it
 * is an expression). Throws for a malformed `Json(...)`.
 */
export function literalOf(text: string): { value: Json } | undefined {
  const t = text.trim()
  if (/^"(?:[^"]|"")*"$/.test(t)) return { value: t.slice(1, -1).replace(/""/g, '"') }
  if (NUMBER.test(t)) {
    const n = Number(t)
    if (Number.isFinite(n)) return { value: n }
  }
  const word = t.toLowerCase()
  if (word === 'true') return { value: true }
  if (word === 'false') return { value: false }
  if (word === 'null') return { value: null }
  const json = /^json\s*\(/i.exec(t)
  if (json) {
    if (closingParen(t, json[0].length - 1) !== t.length - 1)
      throw new Error('bad Json(...): the brackets do not match')
    const inner = t.slice(json[0].length, -1)
    try {
      return { value: JSON.parse(inner) as Json }
    } catch (error) {
      throw new Error(`bad Json(...): ${error instanceof Error ? error.message : String(error)}`, {
        cause: error,
      })
    }
  }
  return undefined
}

/**
 * Expressions in scripts use double-quoted text only (a `'` starts a
 * comment): rewrites 'single-quoted' text of an expression to "double".
 */
export function doubleQuoted(expr: string): string {
  let out = ''
  let i = 0
  while (i < expr.length) {
    const ch = expr[i]!
    if (ch === '"') {
      let j = i + 1
      for (;;) {
        if (j >= expr.length) return out + expr.slice(i)
        if (expr[j] === '"') {
          if (expr[j + 1] === '"') {
            j += 2
            continue
          }
          break
        }
        j++
      }
      out += expr.slice(i, j + 1)
      i = j + 1
      continue
    }
    if (ch === "'") {
      let j = i + 1
      let text = ''
      for (;;) {
        if (j >= expr.length) return out + expr.slice(i)
        if (expr[j] === "'") {
          if (expr[j + 1] === "'") {
            text += "'"
            j += 2
            continue
          }
          break
        }
        text += expr[j]
        j++
      }
      out += `"${text.replace(/"/g, '""')}"`
      i = j + 1
      continue
    }
    out += ch
    i++
  }
  return out
}

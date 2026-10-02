import { useEffect, useRef } from 'react'
import { autocompletion, completionKeymap } from '@codemirror/autocomplete'
import { defaultKeymap, history, historyKeymap } from '@codemirror/commands'
import { HighlightStyle, StreamLanguage, syntaxHighlighting } from '@codemirror/language'
import { linter, lintGutter, type Diagnostic as CmDiagnostic } from '@codemirror/lint'
import { EditorState, StateEffect, StateField, type Extension } from '@codemirror/state'
import { Decoration, EditorView, keymap, lineNumbers, type DecorationSet } from '@codemirror/view'
import { tags } from '@lezer/highlight'
import { checkScript, completeAt, type ScriptCatalogEntry } from '@genoffice/dvh-script'

const KEYWORDS =
  /^(?:workflow|id|description|catalog|param|as|if|then|elseif|else|end|for|each|in|where|next|try|catch|transaction|log|on|and|or|not|true|false|null|json)\b/i

/** DVH-Script highlighting: comments, text, numbers, keywords and Group.Verb actions. */
const dvhScriptLanguage = StreamLanguage.define<null>({
  token(stream) {
    if (stream.eatSpace()) return null
    if (stream.peek() === "'") {
      stream.skipToEnd()
      return 'comment'
    }
    if (stream.peek() === '"') {
      stream.next()
      while (!stream.eol()) {
        if (stream.next() === '"') {
          if (stream.peek() === '"') stream.next()
          else break
        }
      }
      return 'string'
    }
    if (stream.match(/^-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/)) return 'number'
    if (stream.match(/^[A-Z][A-Za-z]*\.[A-Z][A-Za-z0-9]*(?=\s*\()/)) return 'function'
    if (stream.match(KEYWORDS)) return 'keyword'
    if (stream.match(/^[\p{L}_][\p{L}\p{N}_.]*/u)) return 'variableName'
    stream.next()
    return null
  },
})

// chrome colors only: the editor is UI, not document content
const highlight = HighlightStyle.define([
  { tag: tags.keyword, color: 'var(--accent)', fontWeight: '600' },
  { tag: tags.comment, color: 'var(--text-muted)', fontStyle: 'italic' },
  { tag: tags.string, color: 'var(--success, var(--text))' },
  { tag: tags.number, color: 'var(--accent-dark, var(--accent))' },
  { tag: tags.function(tags.variableName), color: 'var(--accent)' },
])

const theme = EditorView.theme({
  '&': {
    color: 'var(--text)',
    backgroundColor: 'var(--surface)',
    border: '1px solid var(--border)',
    fontSize: '12px',
    maxHeight: '320px',
  },
  '.cm-scroller': { fontFamily: 'var(--font-mono, monospace)', overflow: 'auto' },
  '.cm-gutters': {
    backgroundColor: 'var(--surface)',
    color: 'var(--text-muted)',
    borderRight: '1px solid var(--border)',
  },
  '.cm-content': { caretColor: 'var(--text)' },
  '.cm-activeLine': { backgroundColor: 'var(--hover)' },
  '.dvh-script-paused': { backgroundColor: 'var(--accent-soft)' },
})

const setPausedLine = StateEffect.define<number | null>()
const pausedLineField = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update(value, tr) {
    for (const effect of tr.effects) {
      if (!effect.is(setPausedLine)) continue
      const line = effect.value
      if (line === null || line < 1 || line > tr.state.doc.lines) return Decoration.none
      const at = tr.state.doc.line(line).from
      return Decoration.set([Decoration.line({ class: 'dvh-script-paused' }).range(at)])
    }
    return value.map(tr.changes)
  },
  provide: (field) => EditorView.decorations.from(field),
})

/**
 * The DVH-Script editor (P9): completions and checks come from the action
 * catalog, errors are underlined with their message, and the debugger's
 * paused step is the highlighted line.
 */
export function DvhScriptEditor({
  value,
  onChange,
  catalog,
  pausedLine,
  label,
}: {
  value: string
  onChange: (text: string) => void
  catalog: readonly ScriptCatalogEntry[]
  pausedLine: number | null
  label: string
}) {
  const host = useRef<HTMLDivElement>(null)
  const view = useRef<EditorView | null>(null)
  const live = useRef({ catalog, onChange })
  live.current = { catalog, onChange }

  useEffect(() => {
    const extensions: Extension[] = [
      lineNumbers(),
      history(),
      lintGutter(),
      keymap.of([...completionKeymap, ...defaultKeymap, ...historyKeymap]),
      dvhScriptLanguage,
      syntaxHighlighting(highlight),
      theme,
      pausedLineField,
      EditorView.lineWrapping,
      EditorView.contentAttributes.of({ 'aria-label': label }),
      autocompletion({
        override: [
          (ctx) => {
            const result = completeAt(ctx.state.doc.toString(), ctx.pos, live.current.catalog)
            if (!result || (result.from === ctx.pos && !ctx.explicit)) return null
            return {
              from: result.from,
              options: result.options.map((o) => ({
                label: o.label,
                type: o.kind === 'action' || o.kind === 'function' ? 'function' : o.kind,
                ...(o.detail ? { detail: o.detail } : {}),
                ...(o.apply ? { apply: o.apply } : {}),
              })),
            }
          },
        ],
      }),
      linter((v) => {
        const doc = v.state.doc
        return checkScript(doc.toString(), live.current.catalog).diagnostics.map(
          (d): CmDiagnostic => {
            const line = doc.line(Math.min(Math.max(1, d.line), doc.lines))
            const from = Math.min(line.from + d.column - 1, line.to)
            return {
              from,
              to: Math.max(from, line.to),
              severity: d.severity,
              message: d.message,
            }
          },
        )
      }),
      EditorView.updateListener.of((update) => {
        if (update.docChanged) live.current.onChange(update.state.doc.toString())
      }),
    ]
    const created = new EditorView({
      parent: host.current!,
      state: EditorState.create({ doc: value, extensions }),
    })
    view.current = created
    return () => {
      created.destroy()
      view.current = null
    }
    // the editor is created once; later values arrive through the effect below
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // a new script from outside (another workflow, the block editor)
  useEffect(() => {
    const v = view.current
    if (v && v.state.doc.toString() !== value)
      v.dispatch({ changes: { from: 0, to: v.state.doc.length, insert: value } })
  }, [value])

  useEffect(() => {
    const v = view.current
    if (!v) return
    v.dispatch({ effects: setPausedLine.of(pausedLine) })
    if (pausedLine !== null && pausedLine >= 1 && pausedLine <= v.state.doc.lines)
      v.dispatch({
        effects: EditorView.scrollIntoView(v.state.doc.line(pausedLine).from, { y: 'center' }),
      })
  }, [pausedLine])

  return <div ref={host} className="dvh-script-editor" />
}

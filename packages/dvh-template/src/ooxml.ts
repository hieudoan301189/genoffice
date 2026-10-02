/**
 * Small, string-level OOXML helpers for template filling. The engine edits
 * WordprocessingML as text — the parts are read, the controls located by a
 * balanced scan, and everything it does not touch stays byte for byte as the
 * template had it (Word round-trips it).
 */

export interface XmlSpan {
  readonly start: number
  readonly end: number
  readonly xml: string
}

/**
 * Every outermost `<tag>…</tag>` (also `<tag …>`), balanced against nested
 * elements of the same name; self-closing ones included.
 */
export function elements(xml: string, tag: string, from = 0, to = xml.length): XmlSpan[] {
  const out: XmlSpan[] = []
  const re = new RegExp(`<(/?)${tag.replace(':', '\\:')}(?=[\\s/>])[^>]*?(/?)>`, 'g')
  re.lastIndex = from
  let depth = 0
  let start = -1
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null && m.index < to) {
    const closing = m[1] === '/'
    const selfClosing = m[2] === '/'
    if (!closing && selfClosing) {
      if (depth === 0) out.push({ start: m.index, end: re.lastIndex, xml: m[0] })
      continue
    }
    if (!closing) {
      if (depth === 0) start = m.index
      depth++
    } else {
      depth--
      if (depth === 0 && start >= 0) {
        out.push({ start, end: re.lastIndex, xml: xml.slice(start, re.lastIndex) })
        start = -1
      }
      if (depth < 0) depth = 0
    }
  }
  return out
}

/** The w:tag value of a `<w:sdt>` (its own sdtPr, not a nested control's). */
export function sdtTag(sdtXml: string): string {
  const pr = /^<w:sdt[^>]*>\s*<w:sdtPr>([\s\S]*?)<\/w:sdtPr>/.exec(sdtXml)
  return pr ? (/<w:tag w:val="([^"]*)"/.exec(pr[1]!)?.[1] ?? '') : ''
}

/** The inner XML of a control's `<w:sdtContent>` (outermost). */
export function sdtContent(sdtXml: string): string {
  const open = /<w:sdtContent(?:\s[^>]*)?>/.exec(sdtXml)
  const close = sdtXml.lastIndexOf('</w:sdtContent>')
  if (!open || close < 0) return ''
  return sdtXml.slice(open.index + open[0].length, close)
}

/** The control with its content replaced (sdtPr and wrapper kept). */
export function withSdtContent(sdtXml: string, content: string): string {
  const open = /<w:sdtContent(?:\s[^>]*)?>/.exec(sdtXml)
  const close = sdtXml.lastIndexOf('</w:sdtContent>')
  if (!open || close < 0) return sdtXml
  return sdtXml.slice(0, open.index + open[0].length) + content + sdtXml.slice(close)
}

export function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/**
 * Runs that show `text` with the formatting of the first run of `runsXml`
 * (a control's content): one run, line breaks as `<w:br/>`, tabs as `<w:tab/>`.
 */
export function runsWithText(runsXml: string, text: string): string {
  const firstRun = elements(runsXml, 'w:r')[0]?.xml ?? '<w:r></w:r>'
  const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/.exec(firstRun)?.[0] ?? ''
  const pieces = text.split(/(\n|\t)/).map((piece) => {
    if (piece === '\n') return '<w:br/>'
    if (piece === '\t') return '<w:tab/>'
    return piece ? `<w:t xml:space="preserve">${escapeText(piece)}</w:t>` : ''
  })
  return `<w:r>${rPr}${pieces.join('')}</w:r>`
}

/** Replaces spans (non-overlapping, any order) in one pass. */
export function replaceSpans(
  xml: string,
  spans: readonly { start: number; end: number; with: string }[],
): string {
  let out = ''
  let at = 0
  for (const span of [...spans].sort((a, b) => a.start - b.start)) {
    out += xml.slice(at, span.start) + span.with
    at = span.end
  }
  return out + xml.slice(at)
}

/** A table cell must hold a paragraph: cells a removal left empty get `<w:p/>`. */
export function repairEmptyCells(xml: string): string {
  return xml.replace(
    /(<w:tc>(?:\s*<w:tcPr>[\s\S]*?<\/w:tcPr>)?)(\s*)<\/w:tc>/g,
    '$1$2<w:p/></w:tc>',
  )
}

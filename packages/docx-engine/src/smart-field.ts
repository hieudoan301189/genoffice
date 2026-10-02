/**
 * DVH Smart Fields (spike S2, docs/dvh-architecture-implementation-plan.md
 * decision D2): a run-level content control tagged `dvh:f:<fieldId>`, usually
 * bound to the DVH model part with w:dataBinding. Runs inside such a control
 * carry its w:sdtPr verbatim, so editing the paragraph regenerates the control
 * (tag, binding, id) around whatever text the field holds.
 */

import { attrsOf, findChild, serializeXNode, type XNode } from './xml-utils'

export const DVH_FIELD_TAG_PREFIX = 'dvh:f:'

/** The w:sdtPr XML of a DVH Smart Field control, or null for any other w:sdt. */
export function dvhFieldSdtPr(sdt: XNode): string | null {
  const sdtPr = findChild(sdt, 'w:sdtPr')
  const tag = sdtPr ? findChild(sdtPr, 'w:tag') : undefined
  const value = tag ? attrsOf(tag)['w:val'] : undefined
  if (!sdtPr || typeof value !== 'string' || !value.startsWith(DVH_FIELD_TAG_PREFIX)) return null
  return serializeXNode(sdtPr)
}

/** The field id (`f_…`) a Smart Field's w:sdtPr declares, or null. */
export function dvhFieldId(sdtPrXml: string): string | null {
  const m = /<w:tag\s+w:val="dvh:f:([^"]+)"/.exec(sdtPrXml)
  return m ? m[1]! : null
}

export const DVH_TABLE_TAG_PREFIX = 'dvh:t:'

/**
 * The w:sdtPr of a block content control tagged `dvh:t:<tableId>` (a rendered
 * DVH.Table), sliced from the raw `<w:sdt>` XML; null for any other control.
 */
export function dvhTableSdtPrFromXml(sdtXml: string): string | null {
  const m = /^<w:sdt>\s*(<w:sdtPr>[\s\S]*?<\/w:sdtPr>)/.exec(sdtXml.trim())
  if (!m || !/<w:tag\s+w:val="dvh:t:[^"]+"/.test(m[1]!)) return null
  return m[1]!
}

/** The table id (`t_…`) a DVH table control's w:sdtPr declares, or null. */
export function dvhTableId(sdtPrXml: string): string | null {
  const m = /<w:tag\s+w:val="dvh:t:([^"]+)"/.exec(sdtPrXml)
  return m ? m[1]! : null
}

/** Wraps a generated `<w:tbl>` in the DVH table control again (no-op when already wrapped). */
export function wrapDvhTable(tableXml: string, sdtPrXml: string): string {
  if (tableXml.trimStart().startsWith('<w:sdt')) return tableXml
  return `<w:sdt>${sdtPrXml}<w:sdtContent>${tableXml}</w:sdtContent></w:sdt>`
}

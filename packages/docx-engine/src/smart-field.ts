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

/**
 * Release policy (P5): a file that leaves the user's hands — exported, sent,
 * packaged — chooses what DVH data it carries:
 *
 * - `all`: everything, as saved;
 * - `from`: the history from a point in time on (older change sets dropped);
 * - `none`: no history part; Smart Data (model, bound fields) stays;
 * - `strip`: no DVH at all — model and history parts removed, every DVH
 *   content control unwrapped so its text stays as ordinary content.
 *
 * History can hold values deleted since, so a clean copy is the safe default
 * for sending documents out.
 */
import JSZip from 'jszip'
import {
  HISTORY_NS,
  historyFrom,
  MODEL_NS,
  parseHistoryXml,
  serializeHistoryXml,
} from '@genoffice/dvh-model'
import { findCustomXmlItemByNamespace, itemPropsPathOf } from '@genoffice/docx-engine'

export type DvhReleasePolicy =
  | { readonly kind: 'all' }
  | { readonly kind: 'from'; readonly at: string }
  | { readonly kind: 'none' }
  | { readonly kind: 'strip' }

/** Removes one customXml item with its properties, relationships and content type. */
async function removeCustomXmlItem(zip: JSZip, itemPath: string): Promise<void> {
  const props = itemPropsPathOf(itemPath)
  const rels = itemPath.replace(/^customXml\/(item\d+\.xml)$/, 'customXml/_rels/$1.rels')
  zip.remove(itemPath)
  zip.remove(props)
  zip.remove(rels)
  const target = `../${itemPath}`
  for (const relsPath of Object.keys(zip.files).filter((p) => /_rels\/[^/]+\.rels$/.test(p))) {
    const xml = await zip.file(relsPath)!.async('string')
    const next = xml.replace(/<Relationship\b[^>]*\/>/g, (rel) =>
      rel.includes(`Target="${target}"`) || rel.includes(`Target="/${itemPath}"`) ? '' : rel,
    )
    if (next !== xml) zip.file(relsPath, next)
  }
  const typesPath = '[Content_Types].xml'
  const types = await zip.file(typesPath)?.async('string')
  if (types) {
    zip.file(
      typesPath,
      types.replace(/<Override\b[^>]*\/>/g, (o) => (o.includes(`PartName="/${props}"`) ? '' : o)),
    )
  }
}

/**
 * Unwraps every content control tagged `dvh:f:` / `dvh:t:` (innermost first,
 * so a table control holding field controls unwraps fully): the control goes,
 * its content stays.
 */
/** marks a kept control during the scan; cannot occur in XML text (a control character) */
const KEEP = String.fromCharCode(1)

export function unwrapDvhControls(xml: string): string {
  let out = xml
  for (let guard = 0; guard < 10_000; guard++) {
    // an sdt with no nested sdt inside it, whose tag is a DVH one
    const m = /<w:sdt>((?:(?!<w:sdt>)[\s\S])*?)<\/w:sdt>/.exec(out)
    if (!m) break
    const inner = m[1]!
    const isDvh = /<w:tag w:val="dvh:[ft]:/.test(inner)
    const content = /<w:sdtContent>([\s\S]*)<\/w:sdtContent>/.exec(inner)?.[1] ?? ''
    // other controls are kept: mark them so the scan moves past them, restored below
    const replacement = isDvh ? content : `<w:sdt${KEEP}>${inner}</w:sdt${KEEP}>`
    out = out.slice(0, m.index) + replacement + out.slice(m.index + m[0].length)
  }
  return out.replaceAll(`<w:sdt${KEEP}>`, '<w:sdt>').replaceAll(`</w:sdt${KEEP}>`, '</w:sdt>')
}

/** Applies a release policy to a docx or xlsx package; returns the new bytes. */
export async function applyReleasePolicy(
  bytes: Uint8Array,
  policy: DvhReleasePolicy,
): Promise<Uint8Array> {
  if (policy.kind === 'all') return bytes
  const zip = await JSZip.loadAsync(bytes)
  const historyPath = await findCustomXmlItemByNamespace(zip, HISTORY_NS)
  if (policy.kind === 'from') {
    if (historyPath) {
      const history = parseHistoryXml(await zip.file(historyPath)!.async('string'))
      zip.file(historyPath, serializeHistoryXml(historyFrom(history, policy.at)))
    }
  } else {
    if (historyPath) await removeCustomXmlItem(zip, historyPath)
    if (policy.kind === 'strip') {
      const modelPath = await findCustomXmlItemByNamespace(zip, MODEL_NS)
      if (modelPath) await removeCustomXmlItem(zip, modelPath)
      for (const path of Object.keys(zip.files)) {
        if (!/^word\/(document|header\d*|footer\d*)\.xml$/.test(path)) continue
        const xml = await zip.file(path)!.async('string')
        const next = unwrapDvhControls(xml)
        if (next !== xml) zip.file(path, next)
      }
      // workbooks: the hidden _dvh.* names go with the model
      const workbook = await zip.file('xl/workbook.xml')?.async('string')
      if (workbook) {
        zip.file(
          'xl/workbook.xml',
          workbook
            .replace(/<definedName\b[^>]*name="_dvh\.[^"]*"[^>]*>[^<]*<\/definedName>/g, '')
            .replace(/<definedNames>\s*<\/definedNames>/, ''),
        )
      }
    }
  }
  // the customXml folder entry goes with its last item
  if (!Object.keys(zip.files).some((p) => p.startsWith('customXml/') && !zip.files[p]!.dir)) {
    for (const p of Object.keys(zip.files).filter((p) => p.startsWith('customXml/'))) zip.remove(p)
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/**
 * Namespaced customXml data parts (the DVH model and history parts, ADR D1/D9).
 * Writers renumber `customXml/itemN.xml` freely (Excel does on every save), so a
 * part is identified by its root element's namespace, never by its file name.
 */

import JSZip from 'jszip'

export interface CustomXmlPart {
  readonly path: string
  readonly xml: string
  /** ds:itemID of the part's itemProps; content controls bind through it */
  readonly storeItemId: string | null
}

export interface CustomXmlPartWrite {
  /** root element namespace that identifies the part */
  readonly ns: string
  readonly xml: string
  /** ds:itemID written to a new part's itemProps (existing parts keep theirs) */
  readonly storeItemId: string
}

const ITEM_PATH = /^customXml\/item\d+\.xml$/

function rootNamespaceMatches(xml: string, ns: string): boolean {
  const root = /<(?:[\w-]+:)?[\w-]+\b[^>]*>/.exec(xml.replace(/^\uFEFF?<\?xml[^>]*\?>\s*/, ''))
  return root !== null && root[0].includes(`"${ns}"`)
}

export async function findCustomXmlItemByNamespace(zip: JSZip, ns: string): Promise<string | null> {
  for (const path of Object.keys(zip.files)) {
    if (!ITEM_PATH.test(path)) continue
    if (rootNamespaceMatches(await zip.file(path)!.async('string'), ns)) return path
  }
  return null
}

export const itemPropsPathOf = (itemPath: string): string =>
  itemPath.replace(/item(\d+)\.xml$/, 'itemProps$1.xml')

/** Reads the customXml part whose root lives in `ns`, or null. */
export async function readCustomXmlPart(
  bytes: Uint8Array,
  ns: string,
): Promise<CustomXmlPart | null> {
  const zip = await JSZip.loadAsync(bytes)
  const path = await findCustomXmlItemByNamespace(zip, ns)
  if (!path) return null
  const props = await zip.file(itemPropsPathOf(path))?.async('string')
  return {
    path,
    xml: await zip.file(path)!.async('string'),
    storeItemId: props ? (/itemID="(\{[0-9A-Fa-f-]+\})"/.exec(props)?.[1] ?? null) : null,
  }
}

export function customXmlItemPropsXml(storeItemId: string, ns: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<ds:datastoreItem ds:itemID="${storeItemId}" ` +
    'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
    `<ds:schemaRefs><ds:schemaRef ds:uri="${ns}"/></ds:schemaRefs></ds:datastoreItem>`
  )
}

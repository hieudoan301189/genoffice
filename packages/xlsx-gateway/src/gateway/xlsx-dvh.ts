/// DVH Smart Data in xlsx (ADR D1, D3): the hidden `_dvh.*` binding names and
/// the namespaced customXml parts (model, history).
///
/// When the editor models the binding names, the save treats its copy as the
/// truth: the file's own `_dvh.*` names are stripped before any structural
/// shift (so deleting a bound cell can no longer abort the save, and a bound
/// sheet can be deleted) and the editor's names are written last, already in
/// post-edit coordinates.

import { withFutureFunctionMarkers } from './future-functions'
import { nextFreeRelationshipId } from './xlsx-sheets'

export const DVH_NAME_PREFIX = '_dvh.'

export interface DvhDefinedName {
  readonly name: string
  readonly formula: string
}

export interface DvhCustomXmlPart {
  /** root element namespace that identifies the part */
  readonly ns: string
  readonly xml: string
  /** ds:itemID for a new part's itemProps (an existing part keeps its own) */
  readonly storeItemId: string
}

export interface DvhWorkbookState {
  /** every `_dvh.*` name as the editor holds it after this session; replaces the file's */
  readonly names: readonly DvhDefinedName[]
  readonly customXmlParts: readonly DvhCustomXmlPart[]
}

const escapeText = (s: string) =>
  s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
const escapeAttr = (s: string) => escapeText(s).replaceAll('"', '&quot;')

const isDvhElement = (element: string) =>
  new RegExp(`\\bname="${DVH_NAME_PREFIX.replace('.', '\\.')}`).test(element)

/// Removes every `_dvh.*` definedName, then (when given) writes `names` as
/// hidden workbook-scope names. An emptied `<definedNames>` section is dropped.
export function applyDvhDefinedNames(
  workbookXml: string,
  names: readonly DvhDefinedName[],
): string {
  for (const entry of names) {
    if (!entry.name.startsWith(DVH_NAME_PREFIX)) {
      throw new Error(`"${entry.name}" is not a DVH binding name.`)
    }
  }
  let xml = workbookXml.replace(
    /<definedName\b[^>]*>[\s\S]*?<\/definedName>|<definedName\b[^>]*\/>/g,
    (element) => (isDvhElement(element) ? '' : element),
  )
  const additions = names
    .map(
      (entry) =>
        `<definedName name="${escapeAttr(entry.name)}" hidden="1">` +
        `${escapeText(withFutureFunctionMarkers(entry.formula.replace(/^=/, '')))}</definedName>`,
    )
    .join('')
  const section = /<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>|<definedNames\b[^>]*\/>/.exec(
    xml,
  )
  if (section) {
    const inner = (section[1] ?? '') + additions
    const replacement = inner === '' ? '' : `<definedNames>${inner}</definedNames>`
    xml = xml.slice(0, section.index) + replacement + xml.slice(section.index + section[0].length)
  } else if (additions !== '') {
    // CT_Workbook order: ... sheets, functionGroups, externalReferences, definedNames, calcPr ...
    const anchor = xml.includes('</externalReferences>') ? '</externalReferences>' : '</sheets>'
    xml = xml.replace(anchor, `${anchor}<definedNames>${additions}</definedNames>`)
  }
  return xml
}

/// The subset of the gateway's package editor the customXml writer needs.
export interface DvhPackage {
  paths(): Promise<readonly string[]>
  has(path: string): Promise<boolean>
  readText(path: string): Promise<string>
  write(path: string, content: string): void
  add(path: string, content: string): void
}

const ITEM_PATH = /^customXml\/item(\d+)\.xml$/
const CUSTOM_XML_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
const CUSTOM_XML_PROPS_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
const CUSTOM_XML_PROPS_TYPE =
  'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'

function rootNamespaceMatches(xml: string, ns: string): boolean {
  const root = /<(?:[\w-]+:)?[\w-]+\b[^>]*>/.exec(xml.replace(/^\uFEFF?<\?xml[^>]*\?>\s*/, ''))
  return root !== null && root[0].includes(`"${ns}"`)
}

/// Rewrites each part in place when a part with its namespace exists (writers
/// renumber items freely), otherwise adds item + itemProps + relationships.
export async function applyDvhCustomXmlParts(
  pkg: DvhPackage,
  parts: readonly DvhCustomXmlPart[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (parts.length === 0) return
  const items = (await pkg.paths()).filter((path) => ITEM_PATH.test(path))
  const byNamespace = new Map<string, string>()
  for (const path of items) {
    const xml = await pkg.readText(path)
    for (const part of parts) {
      if (!byNamespace.has(part.ns) && rootNamespaceMatches(xml, part.ns))
        byNamespace.set(part.ns, path)
    }
  }
  const used = new Set(items.map((path) => Number(ITEM_PATH.exec(path)![1])))
  for (const part of parts) {
    const existing = byNamespace.get(part.ns)
    if (existing) {
      pkg.write(existing, part.xml)
      touchedEntries.add(existing)
      continue
    }
    let n = 1
    while (used.has(n)) n++
    used.add(n)
    const itemPath = `customXml/item${n}.xml`
    const propsPath = `customXml/itemProps${n}.xml`
    pkg.add(itemPath, part.xml)
    pkg.add(
      propsPath,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        `<ds:datastoreItem ds:itemID="${part.storeItemId}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">` +
        `<ds:schemaRefs><ds:schemaRef ds:uri="${escapeAttr(part.ns)}"/></ds:schemaRefs></ds:datastoreItem>`,
    )
    pkg.add(
      `customXml/_rels/item${n}.xml.rels`,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="${CUSTOM_XML_PROPS_REL}" Target="itemProps${n}.xml"/></Relationships>`,
    )
    touchedEntries.add(itemPath).add(propsPath).add(`customXml/_rels/item${n}.xml.rels`)

    const relsPath = 'xl/_rels/workbook.xml.rels'
    const rels = await pkg.readText(relsPath)
    pkg.write(
      relsPath,
      rels.replace(
        '</Relationships>',
        `<Relationship Id="${nextFreeRelationshipId(rels)}" Type="${CUSTOM_XML_REL}" Target="../${itemPath}"/></Relationships>`,
      ),
    )
    touchedEntries.add(relsPath)
    const typesPath = '[Content_Types].xml'
    const types = await pkg.readText(typesPath)
    if (!types.includes(`PartName="/${propsPath}"`)) {
      pkg.write(
        typesPath,
        types.replace(
          '</Types>',
          `<Override PartName="/${propsPath}" ContentType="${CUSTOM_XML_PROPS_TYPE}"/></Types>`,
        ),
      )
      touchedEntries.add(typesPath)
    }
  }
}

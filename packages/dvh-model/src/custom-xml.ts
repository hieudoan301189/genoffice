/**
 * The model part (ADR D1): `customXml/itemN.xml` rooted at
 * `<dvh:model xmlns:dvh="urn:dvh-office:model:1">`. Fields are XML elements so
 * a content control's w:dataBinding XPath can reach them; collections and links
 * travel as JSON in CDATA. Word and Excel re-serialize the part (attribute
 * order, CDATA, part numbering), so reading is tolerant and parts are found by
 * namespace, never by file name.
 */

import {
  dvhModelSchema,
  fieldValueFromText,
  fieldText,
  type DvhModel,
  type FieldType,
} from './types'

export const MODEL_NS = 'urn:dvh-office:model:1'
export const HISTORY_NS = 'urn:dvh-office:history:1'
export const CUSTOM_XML_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
export const CUSTOM_XML_PROPS_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
export const CUSTOM_XML_PROPS_TYPE =
  'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'
/** w:dataBinding prefix mapping for the model namespace */
export const MODEL_PREFIX_MAPPINGS = `xmlns:dvh='${MODEL_NS}'`

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

export function decodeXml(text: string): string {
  return text
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1')
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(Number(dec)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** XPath a content control binds to for one field (attribute predicate, verified in Word by spike S1). */
export function fieldXPath(fieldId: string): string {
  return `/dvh:model[1]/dvh:fields[1]/dvh:f[@id='${fieldId}'][1]`
}

const cdata = (text: string) => `<![CDATA[${text.replace(/\]\]>/g, ']]]]><![CDATA[>')}]]>`

export function serializeModelXml(model: DvhModel): string {
  const fields = model.fields
    .map((f) => {
      const enumAttr = f.enumValues ? ` enum="${escapeXml(JSON.stringify(f.enumValues))}"` : ''
      const revAttr = f.rev !== undefined ? ` rev="${f.rev}"` : ''
      return (
        `<dvh:f id="${escapeXml(f.id)}" name="${escapeXml(f.name)}" type="${f.type}" access="${f.access}"${enumAttr}${revAttr}>` +
        `${escapeXml(fieldText(f.value))}</dvh:f>`
      )
    })
    .join('')
  const objects = JSON.stringify({
    collections: model.collections,
    tables: model.tables,
    links: model.links,
    ...(model.conditions ? { conditions: model.conditions } : {}),
  })
  return (
    XML_DECL +
    `<dvh:model xmlns:dvh="${MODEL_NS}" docId="${escapeXml(model.docId)}" schemaVersion="${model.schemaVersion}">` +
    `<dvh:fields>${fields}</dvh:fields>` +
    `<dvh:objects>${cdata(objects)}</dvh:objects>` +
    '</dvh:model>'
  )
}

function attr(tag: string, name: string): string | undefined {
  const m =
    new RegExp(`\\s${name}="([^"]*)"`).exec(tag) ?? new RegExp(`\\s${name}='([^']*)'`).exec(tag)
  return m ? decodeXml(m[1]!) : undefined
}

/** Parses a model part written by DVH Office, Word or Excel; throws on anything that is not a valid model. */
export function parseModelXml(xml: string): DvhModel {
  const root = /<(\w+:)?model\b[^>]*>/.exec(xml)
  if (!root || !xml.includes(MODEL_NS)) throw new Error('not a DVH model part')
  const fields = []
  for (const m of xml.matchAll(/<(?:\w+:)?f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?f>)/g)) {
    const tag = `<f${m[1] ?? ''}>`
    const type = (attr(tag, 'type') ?? 'text') as FieldType
    const enumText = attr(tag, 'enum')
    const rev = Number(attr(tag, 'rev'))
    fields.push({
      id: attr(tag, 'id') ?? '',
      name: attr(tag, 'name') ?? '',
      type,
      access: attr(tag, 'access') ?? 'readwrite',
      value: fieldValueFromText(type, decodeXml(m[2] ?? '')),
      ...(enumText ? { enumValues: JSON.parse(enumText) as string[] } : {}),
      ...(Number.isInteger(rev) && rev >= 0 ? { rev } : {}),
    })
  }
  const objectsMatch = /<(?:\w+:)?objects\b[^>]*>([\s\S]*?)<\/(?:\w+:)?objects>/.exec(xml)
  const objects = objectsMatch
    ? (JSON.parse(decodeXml(objectsMatch[1]!)) as Record<string, unknown>)
    : {}
  return dvhModelSchema.parse({
    docId: attr(root[0], 'docId'),
    schemaVersion: Number(attr(root[0], 'schemaVersion') ?? 1),
    fields,
    collections: objects.collections ?? [],
    tables: objects.tables ?? [],
    links: objects.links ?? [],
    ...(objects.conditions ? { conditions: objects.conditions } : {}),
  })
}

/**
 * Replaces one field's text inside an existing model part, leaving every other
 * byte alone (what Word's own two-way binding does); `rev` also sets the
 * field's revision. Returns null when the part has no such field.
 */
export function setFieldTextInXml(
  xml: string,
  fieldId: string,
  text: string,
  rev?: number,
): string | null {
  const re = new RegExp(
    `(<(?:\\w+:)?f\\b[^>]*\\sid=["']${fieldId.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}["'][^>]*?)(?:/>|>[\\s\\S]*?</((?:\\w+:)?f)>)`,
  )
  const m = re.exec(xml)
  if (!m) return null
  const close = m[2] ?? (/<(\w+:)?f\b/.exec(m[1]!)?.[1] ?? '') + 'f'
  let open = m[1]!
  if (rev !== undefined) {
    open = /\srev=["'][^"']*["']/.test(open)
      ? open.replace(/\srev=["'][^"']*["']/, ` rev="${Math.trunc(rev)}"`)
      : `${open} rev="${Math.trunc(rev)}"`
  }
  return (
    xml.slice(0, m.index) +
    `${open}>${escapeXml(text)}</${close}>` +
    xml.slice(m.index + m[0].length)
  )
}

export function itemPropsXml(storeItemId: string, ns: string): string {
  return (
    XML_DECL +
    `<ds:datastoreItem ds:itemID="${storeItemId}" ` +
    'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">' +
    `<ds:schemaRefs><ds:schemaRef ds:uri="${ns}"/></ds:schemaRefs></ds:datastoreItem>`
  )
}

export function storeItemIdOf(itemPropsXmlText: string): string | null {
  return /itemID="(\{[0-9A-Fa-f-]+\})"/.exec(itemPropsXmlText)?.[1] ?? null
}

export interface PackagePart {
  readonly path: string
  readonly xml: string
}

/** The customXml item whose root element lives in `ns`, wherever the writer numbered it. */
export function findPartByNamespace(parts: Iterable<PackagePart>, ns: string): PackagePart | null {
  for (const part of parts) {
    if (!/^customXml\/item\d+\.xml$/.test(part.path)) continue
    const root = /<(?:\w+:)?[\w-]+\b[^>]*>/.exec(part.xml.replace(/^<\?xml[^>]*\?>\s*/, ''))
    if (root && root[0].includes(`"${ns}"`)) return part
  }
  return null
}

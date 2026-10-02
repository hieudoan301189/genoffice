/**
 * Where DVH objects surface in documents (ADR D2, D3): a content control tag in
 * docx, a hidden defined name in xlsx.
 */

import { escapeXml, fieldXPath, MODEL_PREFIX_MAPPINGS } from './custom-xml'

export const FIELD_SDT_TAG_PREFIX = 'dvh:f:'
export const TABLE_SDT_TAG_PREFIX = 'dvh:t:'
export const FIELD_NAME_PREFIX = '_dvh.f.'
export const TABLE_NAME_PREFIX = '_dvh.t.'
/** a collection's source range in a workbook: title row + data rows */
export const COLLECTION_NAME_PREFIX = '_dvh.c.'

export const fieldSdtTag = (fieldId: string): string => `${FIELD_SDT_TAG_PREFIX}${fieldId}`
export const tableSdtTag = (tableId: string): string => `${TABLE_SDT_TAG_PREFIX}${tableId}`
export const fieldDefinedName = (fieldId: string): string => `${FIELD_NAME_PREFIX}${fieldId}`
export const tableDefinedName = (tableId: string): string => `${TABLE_NAME_PREFIX}${tableId}`
export const collectionDefinedName = (collectionId: string): string =>
  `${COLLECTION_NAME_PREFIX}${collectionId}`

/** True for the defined names DVH owns; they are system names, never shown in the Name Manager. */
export function isDvhDefinedName(name: string): boolean {
  return (
    name.startsWith(FIELD_NAME_PREFIX) ||
    name.startsWith(TABLE_NAME_PREFIX) ||
    name.startsWith(COLLECTION_NAME_PREFIX)
  )
}

export function parseDvhDefinedName(
  name: string,
): { kind: 'field' | 'table' | 'collection'; id: string } | null {
  if (name.startsWith(FIELD_NAME_PREFIX))
    return { kind: 'field', id: name.slice(FIELD_NAME_PREFIX.length) }
  if (name.startsWith(TABLE_NAME_PREFIX))
    return { kind: 'table', id: name.slice(TABLE_NAME_PREFIX.length) }
  if (name.startsWith(COLLECTION_NAME_PREFIX))
    return { kind: 'collection', id: name.slice(COLLECTION_NAME_PREFIX.length) }
  return null
}

export function parseDvhSdtTag(tag: string): { kind: 'field' | 'table'; id: string } | null {
  if (tag.startsWith(FIELD_SDT_TAG_PREFIX))
    return { kind: 'field', id: tag.slice(FIELD_SDT_TAG_PREFIX.length) }
  if (tag.startsWith(TABLE_SDT_TAG_PREFIX))
    return { kind: 'table', id: tag.slice(TABLE_SDT_TAG_PREFIX.length) }
  return null
}

/**
 * The w:sdtPr of a Smart Field content control: alias, tag, a document-unique
 * numeric id, and the data binding Word resolves against the model part.
 */
export function fieldSdtPrXml(options: {
  fieldId: string
  alias: string
  sdtId: number
  storeItemId: string
}): string {
  return (
    '<w:sdtPr>' +
    `<w:alias w:val="${escapeXml(options.alias)}"/>` +
    `<w:tag w:val="${fieldSdtTag(options.fieldId)}"/>` +
    `<w:id w:val="${Math.trunc(options.sdtId)}"/>` +
    `<w:dataBinding w:prefixMappings="${MODEL_PREFIX_MAPPINGS}" ` +
    `w:xpath="${fieldXPath(options.fieldId)}" w:storeItemID="${options.storeItemId}"/>` +
    '<w:text/></w:sdtPr>'
  )
}

/**
 * The w:sdtPr of the block content control that wraps a rendered DVH.Table.
 * A table has no data binding (Word cannot bind rows); the tag tells DVH
 * Office which table to refresh.
 */
export function tableSdtPrXml(options: { tableId: string; alias: string; sdtId: number }): string {
  return (
    '<w:sdtPr>' +
    `<w:alias w:val="${escapeXml(options.alias)}"/>` +
    `<w:tag w:val="${tableSdtTag(options.tableId)}"/>` +
    `<w:id w:val="${Math.trunc(options.sdtId)}"/>` +
    '</w:sdtPr>'
  )
}

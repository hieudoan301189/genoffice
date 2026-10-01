/**
 * Stable DVH identifiers (ADR D4): `<prefix>_<16 base32 chars>`. An id never
 * changes; names like `Project.Name` are aliases on top of it.
 */

export type DvhIdPrefix = 'f' | 'r' | 'c' | 't' | 'l' | 'wf' | 'doc'

const ALPHABET = 'abcdefghijklmnopqrstuvwxyz234567'
const ID_BODY = 16
const ID_PATTERN = /^(f|r|c|t|l|wf|doc)_[a-z2-7]{16}$/

function randomBytes(count: number): Uint8Array {
  const bytes = new Uint8Array(count)
  globalThis.crypto.getRandomValues(bytes)
  return bytes
}

export function newDvhId(prefix: DvhIdPrefix): string {
  let body = ''
  for (const byte of randomBytes(ID_BODY)) body += ALPHABET[byte & 31]
  return `${prefix}_${body}`
}

export function isDvhId(value: unknown, prefix?: DvhIdPrefix): value is string {
  if (typeof value !== 'string') return false
  const m = ID_PATTERN.exec(value)
  return m !== null && (prefix === undefined || m[1] === prefix)
}

/** Store item id of a customXml part (`{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`); Word binds content controls through it. */
export function newStoreItemId(): string {
  return `{${globalThis.crypto.randomUUID().toUpperCase()}}`
}

import { describe, expect, it } from 'vitest'
import {
  checkFieldWrite,
  fieldBaseOf,
  parseModelXml,
  serializeModelXml,
  setFieldTextInXml,
  threeWay,
  writeFieldsRequestSchema,
  type DvhModel,
} from '../src/index'

const model: DvhModel = {
  docId: 'doc_aaaaaaaaaaaaaaaa',
  schemaVersion: 1,
  fields: [
    { id: 'f_a', name: 'Project.Name', type: 'text', value: 'A', access: 'readwrite', rev: 3 },
    { id: 'f_b', name: 'Lot.Count', type: 'number', value: 12, access: 'readwrite' },
  ],
  collections: [],
  tables: [],
  links: [],
}

describe('field revisions', () => {
  it('round-trips rev through the model part and leaves fields without one alone', () => {
    const parsed = parseModelXml(serializeModelXml(model))
    expect(parsed.fields[0]!.rev).toBe(3)
    expect(parsed.fields[1]!.rev).toBeUndefined()
  })

  it('setFieldTextInXml sets or replaces rev without touching other fields', () => {
    const xml = serializeModelXml(model)
    const once = setFieldTextInXml(xml, 'f_a', 'B & C', 4)!
    const twice = setFieldTextInXml(once, 'f_b', '13', 1)!
    const parsed = parseModelXml(twice)
    expect(parsed.fields[0]).toMatchObject({ value: 'B & C', rev: 4 })
    expect(parsed.fields[1]).toMatchObject({ value: 13, rev: 1 })
    expect(twice.match(/rev=/g)).toHaveLength(2)
  })
})

describe('write-back conflicts', () => {
  const base = { rev: 3, text: 'A' }

  it('three-way: local only, source only, both, agreed', () => {
    expect(threeWay(base, 'B', { rev: 3, text: 'A' })).toBe('local')
    expect(threeWay(base, 'A', { rev: 4, text: 'C' })).toBe('source')
    expect(threeWay(base, 'B', { rev: 4, text: 'C' })).toBe('conflict')
    expect(threeWay(base, 'C', { rev: 4, text: 'C' })).toBe('same')
    // an Excel edit keeps the revision; the text still tells
    expect(threeWay(base, 'B', { rev: 3, text: 'X' })).toBe('conflict')
    // links made before P3 have no base: the source wins
    expect(threeWay(undefined, 'B', { rev: 0, text: 'X' })).toBe('source')
  })

  it('writer re-check: stale base conflicts unless forced or equal', () => {
    const write = { fieldId: 'f_a', value: 'B', expected: base, force: false }
    expect(checkFieldWrite(write, { rev: 3, text: 'A' })).toBe('write')
    expect(checkFieldWrite(write, { rev: 4, text: 'A' })).toBe('conflict')
    expect(checkFieldWrite({ ...write, force: true }, { rev: 4, text: 'Z' })).toBe('write')
    expect(checkFieldWrite(write, { rev: 9, text: 'B' })).toBe('write')
    expect(fieldBaseOf(model.fields[1]!)).toEqual({ rev: 0, text: '12' })
  })

  it('validates requests', () => {
    expect(() =>
      writeFieldsRequestSchema.parse({
        docId: 'doc_x',
        paths: [],
        writes: [],
        origin: { docId: 'doc_y' },
      }),
    ).toThrow()
  })
})

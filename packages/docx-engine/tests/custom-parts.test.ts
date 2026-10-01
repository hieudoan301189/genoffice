import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  buildBlankDocx,
  parseDocx,
  readCustomXmlPart,
  saveDocx,
  type SaveBlock,
} from '../src/index'

const MODEL_NS = 'urn:dvh-office:model:1'
const HISTORY_NS = 'urn:dvh-office:history:1'
const MODEL_ID = '{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A01}'
const HISTORY_ID = '{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A02}'
const model = (value: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><dvh:model xmlns:dvh="${MODEL_NS}" docId="doc_x" schemaVersion="1"><dvh:fields><dvh:f id="f_a" name="A" type="text" access="readwrite">${value}</dvh:f></dvh:fields><dvh:objects><![CDATA[{"collections":[],"links":[]}]]></dvh:objects></dvh:model>`
const history = `<dvh:history xmlns:dvh="${HISTORY_NS}" docId="doc_x" count="0"><dvh:changes><![CDATA[]]></dvh:changes></dvh:history>`

async function saveWith(
  bytes: Uint8Array,
  parts: { ns: string; xml: string; storeItemId: string }[],
) {
  const doc = await parseDocx(bytes)
  const blocks: SaveBlock[] = doc.blocks
    .filter((b) => !b.hidden)
    .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
  return saveDocx(doc, blocks, { customXmlParts: parts })
}

describe('namespaced customXml parts', () => {
  it('adds new parts with itemProps, relationships and content types', async () => {
    const saved = await saveWith(await buildBlankDocx(), [
      { ns: MODEL_NS, xml: model('Dự án A'), storeItemId: MODEL_ID },
      { ns: HISTORY_NS, xml: history, storeItemId: HISTORY_ID },
    ])
    const zip = await JSZip.loadAsync(saved)
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    const types = await zip.file('[Content_Types].xml')!.async('string')
    expect(rels.match(/relationships\/customXml"/g)).toHaveLength(2)
    expect(types).toContain('PartName="/customXml/itemProps1.xml"')
    expect(types).toContain('PartName="/customXml/itemProps2.xml"')
    expect(await zip.file('customXml/_rels/item1.xml.rels')!.async('string')).toContain(
      'Target="itemProps1.xml"',
    )
    const part = await readCustomXmlPart(saved, MODEL_NS)
    expect(part).toMatchObject({ path: 'customXml/item1.xml', storeItemId: MODEL_ID })
    expect(part!.xml).toContain('Dự án A')
    expect((await readCustomXmlPart(saved, HISTORY_NS))?.storeItemId).toBe(HISTORY_ID)
    // the package still parses
    expect((await parseDocx(saved)).blocks.length).toBeGreaterThan(0)
  })

  it('rewrites an existing part in place and keeps its store id', async () => {
    const first = await saveWith(await buildBlankDocx(), [
      { ns: MODEL_NS, xml: model('A'), storeItemId: MODEL_ID },
    ])
    const second = await saveWith(first, [
      { ns: MODEL_NS, xml: model('B'), storeItemId: '{00000000-0000-0000-0000-000000000000}' },
    ])
    const zip = await JSZip.loadAsync(second)
    expect(Object.keys(zip.files).filter((p) => /^customXml\/item\d+\.xml$/.test(p))).toEqual([
      'customXml/item1.xml',
    ])
    expect(
      (await zip.file('word/_rels/document.xml.rels')!.async('string')).match(
        /relationships\/customXml"/g,
      ),
    ).toHaveLength(1)
    const part = await readCustomXmlPart(second, MODEL_NS)
    expect(part?.xml).toContain('>B<')
    expect(part?.storeItemId).toBe(MODEL_ID)
  })

  it('finds a part by namespace when another writer numbered it differently', async () => {
    const zip = await JSZip.loadAsync(await buildBlankDocx())
    zip.file('customXml/item7.xml', model('Excel-numbered'))
    const bytes = new Uint8Array(await zip.generateAsync({ type: 'uint8array' }))
    expect((await readCustomXmlPart(bytes, MODEL_NS))?.path).toBe('customXml/item7.xml')
    expect(await readCustomXmlPart(bytes, HISTORY_NS)).toBeNull()
  })
})

/**
 * P6: a Smart Template opened in Docs keeps its controls — column values and
 * image fields at run level (kept like Smart Fields), conditions and
 * repeating sections at block level (content-control shells).
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'

const X = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const sdt = (tag: string, content: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`

async function template(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${X}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `${X}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  const body =
    `<w:p>${run('Mã: ')}${sdt('dvh:c:k_code', run('mã'))}</w:p>` +
    sdt('dvh:if:cond_fail', `<w:p>${run('KHÔNG ĐẠT')}</w:p>`) +
    sdt(
      'dvh:repeat:c_checks',
      `<w:p>${run('Lỗi: ')}${sdt('dvh:c:k_item', run('?'))}</w:p><w:p>${run('---')}</w:p>`,
    )
  zip.file(
    'word/document.xml',
    `${X}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'uint8array' })
}

describe('Smart Template in Docs', () => {
  it('keeps column controls on their runs and section controls on their blocks', async () => {
    const parsed = await parseDocx(await template())
    const blocks = parsed.blocks
    const runs = blocks.flatMap((b) => ('runs' in b ? (b.runs ?? []) : []))
    expect(runs.filter((r) => r.sdtFieldXml?.includes('dvh:c:k_code')).map((r) => r.text)).toEqual([
      'mã',
    ])
    expect(runs.filter((r) => r.sdtFieldXml?.includes('dvh:c:k_item')).map((r) => r.text)).toEqual([
      '?',
    ])
    const tags = blocks.map((b) => b.sdtShell?.tag ?? '')
    // (the parser adds a trailing empty block for the section break)
    expect(tags.slice(0, 4)).toEqual([
      '',
      'dvh:if:cond_fail',
      'dvh:repeat:c_checks',
      'dvh:repeat:c_checks',
    ])
    expect(tags.slice(4).every((t) => t === '')).toBe(true)
    // one shell spans the two paragraphs of the repeated section
    const repeat = blocks.filter((b) => b.sdtShell?.tag === 'dvh:repeat:c_checks')
    expect(repeat[0]!.sdtShell!.openXml).toContain('dvh:repeat:c_checks')
    expect(repeat[1]!.sdtShell!.closeXml).toContain('</w:sdt>')
  })
})

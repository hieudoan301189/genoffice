/**
 * P6 Smart Template engine. M2 acceptance: 100 acceptance records generated
 * from WorkItems; a second run gives identical bytes; the output stays a valid
 * package (Word opens it: checked on Windows, see docs/phases/p6-*.md).
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseModelXml } from '@genoffice/dvh-model'
import {
  batchIndexCsv,
  evaluateExpr,
  exprError,
  fillTemplate,
  generateBatch,
  packageBatch,
  planBatch,
  renderTemplateText,
  safeFileName,
  type ExprScope,
} from '../src/index'
import { templateDocx, workModel } from './fixtures'

const scope = (
  values: Record<string, unknown>,
  columns: Record<string, unknown> = {},
): ExprScope => ({
  lookup: (name) => values[name] as never,
  column: (title) => columns[title] as never,
})

describe('expressions', () => {
  it('evaluates values, comparisons, logic, text and functions', () => {
    const s = scope(
      { 'Project.Name': 'Cầu', 'row.qty': 1250, 'row.ok': false, INDEX: 7 },
      { 'Khối lượng': 3 },
    )
    expect(evaluateExpr('row.qty >= 1100 AND NOT row.ok', s)).toBe(true)
    expect(evaluateExpr('"BB-" & PAD(INDEX, 3) & "-" & UPPER(Project.Name)', s)).toBe('BB-007-CẦU')
    expect(evaluateExpr('ROUND(row.qty / 3, 2)', s)).toBe(416.67)
    expect(evaluateExpr('IF([Khối lượng] > 2, "lớn", 1/0)', s)).toBe('lớn')
    expect(evaluateExpr('=2^3^2', s)).toBe(512)
    expect(evaluateExpr("'it''s' = \"it's\"", s)).toBe(true)
    expect(evaluateExpr('CONTAINS("Bê tông lót", "TÔNG")', s)).toBe(true)
    expect(renderTemplateText('BB-{INDEX}-{{x}}', s)).toBe('BB-7-{x}')
  })

  it('never runs code and reports errors', () => {
    expect(exprError('constructor.constructor("alert(1)")()')).toMatch(/unexpected|unknown/)
    expect(() => evaluateExpr('process.exit', scope({}))).toThrow(/unknown name/)
    expect(() => evaluateExpr('1/0', scope({}))).toThrow(/division by zero/)
    expect(exprError('(1 + 2')).toMatch(/expected/)
    expect(exprError('x'.repeat(3000))).toMatch(/too long/)
  })
})

describe('fillTemplate', () => {
  it('fills fields, columns, conditions, repeated rows/blocks, tables and images', async () => {
    const model = workModel(3)
    const result = await fillTemplate(
      await templateDocx(model),
      {
        model,
        record: { collectionId: 'c_work', index: 0, row: model.collections[0]!.rows[0]! },
        count: 3,
      },
      { locale: 'vi-VN', images: () => ({ bytes: new Uint8Array([9, 9]), ext: 'png' }) },
    )
    expect(result.warnings).toEqual([])
    const zip = await JSZip.loadAsync(result.bytes)
    const xml = await zip.file('word/document.xml')!.async('string')
    // field control kept, value written, model part agrees
    expect(xml).toContain('w:val="dvh:f:f_project"')
    expect(xml).toContain('Cầu Bến Thủy 3')
    // record 1 (AB.1, ok = false): the failure paragraph is in, the "large" note is out
    expect(xml).toContain('AB.1')
    expect(xml).toContain('Hạng mục 1 &amp; &lt;phụ&gt;')
    expect(xml).toContain('1.000,00')
    expect(xml).toContain('KHÔNG ĐẠT')
    expect(xml).not.toContain('khối lượng lớn')
    // repeated table rows: header + 3 rows + trailing row; section wrappers gone
    expect(xml.match(/Kích thước/g)).toHaveLength(2) // repeated row + DVH table
    expect(xml).toContain('Hết bảng')
    expect(xml).not.toMatch(/dvh:(repeat|repeatrows|if|c):/)
    // filtered block repeat: only the failed check
    expect(xml).toContain('Lỗi: ')
    expect(xml.match(/Lỗi: /g)).toHaveLength(1)
    // the DVH table re-rendered from its collection, control kept
    expect(xml).toContain('w:val="dvh:t:t_checks"')
    expect(xml).toContain('<w:tblHeader/>')
    // image swapped through its relationship
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain('Target="media/dvh_f_logo.png"')
    expect(await zip.file('word/media/dvh_f_logo.png')!.async('uint8array')).toEqual(
      new Uint8Array([9, 9]),
    )
    const modelPart = Object.keys(zip.files).find((f) => /^customXml\/item\d+\.xml$/.test(f))!
    expect(parseModelXml(await zip.file(modelPart)!.async('string')).fields[0]!.value).toBe(
      'Cầu Bến Thủy 3',
    )
  })

  it('release options: strip leaves no DVH controls or parts', async () => {
    const model = workModel(2)
    const result = await fillTemplate(
      await templateDocx(model),
      { model, record: { collectionId: 'c_work', index: 1, row: model.collections[0]!.rows[1]! } },
      { release: { kind: 'strip' } },
    )
    const zip = await JSZip.loadAsync(result.bytes)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml).not.toContain('dvh:')
    expect(xml).toContain('Cầu Bến Thủy 3')
    expect(Object.keys(zip.files).some((f) => f.startsWith('customXml/'))).toBe(false)
  })
})

describe('batch generator (M2)', () => {
  it('plans names, filters records and flags duplicates', () => {
    const model = workModel(10)
    const plan = planBatch({
      template: new Uint8Array(),
      model,
      collectionId: 'c_work',
      filter: 'row.qty >= 1050',
      nameRule: 'BB {row.code} {IF(row.ok, "dat", "khong dat")}',
    })
    expect(plan.items.map((i) => i.name)).toEqual([
      'BB AB.5 dat.docx',
      'BB AB.6 dat.docx',
      'BB AB.7 khong dat.docx',
      'BB AB.8 dat.docx',
      'BB AB.9 dat.docx',
      'BB AB.10 khong dat.docx',
    ])
    const dup = planBatch({
      template: new Uint8Array(),
      model,
      collectionId: 'c_work',
      nameRule: 'BB',
    })
    expect(dup.duplicates).toEqual(['BB'])
    expect(dup.items[1]!.name).toBe('BB (2).docx')
    expect(safeFileName('a/b:c*?"<>|. ')).toBe('a_b_c______')
    expect(safeFileName('CON')).toBe('CON_')
  })

  it('generates 100 acceptance records, identical on a second run, packaged with an index', async () => {
    const model = workModel(100)
    const spec = {
      template: await templateDocx(model),
      model,
      collectionId: 'c_work',
      nameRule: 'BB-{PAD(INDEX, 3)}-{row.code}',
      options: { locale: 'vi-VN', images: () => ({ bytes: new Uint8Array([7]), ext: 'png' }) },
    }
    const first = await generateBatch(spec)
    expect(first).toHaveLength(100)
    expect(first[99]!.name).toBe('BB-100-AB.100.docx')
    const second = await generateBatch(spec)
    for (let i = 0; i < first.length; i++) expect(second[i]!.bytes).toEqual(first[i]!.bytes)
    // the 100th record holds its own data
    const xml = await (
      await JSZip.loadAsync(first[99]!.bytes)
    )
      .file('word/document.xml')!
      .async('string')
    expect(xml).toContain('AB.100')
    expect(xml).toContain('Hạng mục 100')
    const zip = await packageBatch(model, first)
    expect(await packageBatch(model, second)).toEqual(zip)
    const pkg = await JSZip.loadAsync(zip)
    expect(Object.keys(pkg.files)).toHaveLength(101)
    const index = await pkg.file('_MucLuc.csv')!.async('string')
    expect(index.startsWith('﻿STT,File,Mã,Công việc,Khối lượng,Đạt')).toBe(true)
    expect(index.split('\r\n')).toHaveLength(102)
    expect(batchIndexCsv(model, [])).toBe('﻿STT,File\r\n')
  }, 60_000)
})

// Spike S5: DVH functions that return formatting, checked in the packaged app.
// Usage: node app-check.mjs <outDir> <path to DVH Office.exe>
import { _electron as electron } from 'playwright-core'
import JSZip from 'jszip'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { execFileSync } from 'node:child_process'

const outDir = resolve(process.argv[2])
const exe = resolve(process.argv[3])
mkdirSync(outDir, { recursive: true })
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const report = {}

// ---------------- probe workbook ----------------
const esc = (s) =>
  String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
const str = (ref, text, s = 0) =>
  `<c r="${ref}"${s ? ` s="${s}"` : ''} t="inlineStr"><is><t>${esc(text)}</t></is></c>`
const num = (ref, v, s = 0) => `<c r="${ref}"${s ? ` s="${s}"` : ''}><v>${v}</v></c>`
const fml = (ref, f, cached = '') =>
  `<c r="${ref}" t="str"><f>${esc(f)}</f><v>${esc(cached)}</v></c>`
// styles: 1 header (bold + blue fill), 2 qty (#,##0.000), 3 yellow fill, 4 yellow fill + qty, 5 red font, 6 red font + qty
const styles =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.000"/></numFmts>' +
  '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>' +
  '<font><sz val="11"/><color rgb="FFFF0000"/><name val="Calibri"/></font></fonts>' +
  '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="7"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
  '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>' +
  '<xf numFmtId="164" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/>' +
  '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
  '<xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
const rows = [
  `<row r="1">${['Mã', 'Công việc', 'KQ', 'KL'].map((t, i) => str(`${'ABCD'[i]}1`, t, 1)).join('')}` +
    `${fml('F1', 'DVH.Table(A1:D5,1,,C1:C5="Đạt")', 'OLD')}</row>`,
  `<row r="2">${str('A2', 'CV01')}${str('B2', 'Đào đất hố móng')}${str('C2', 'Đạt')}${num('D2', 125.5, 2)}</row>`,
  `<row r="3">${str('A3', 'CV02', 3)}${str('B3', 'Bê tông lót', 3)}${str('C3', 'Không đạt', 3)}${num('D3', 18.25, 4)}</row>`,
  `<row r="4">${str('A4', 'CV03', 5)}${str('B4', 'Cốt thép', 5)}${str('C4', 'Đạt', 5)}${num('D4', 2.734, 6)}</row>`,
  `<row r="5">${str('A5', 'CV04')}${str('B5', 'Ván khuôn')}${str('C5', 'Đạt')}${num('D5', 40, 2)}</row>`,
  `<row r="8">${str('A8', 'Đậm nghiêng')}${fml('B8', 'DVH.Font(A8:A9,"BI")', 'OLD')}</row>`,
  `<row r="9">${str('A9', 'Dòng 2')}</row>`,
  `<row r="10">${str('A10', 'Nền vàng')}${fml('B10', 'DVH.FillColor(A10,65535)', 'OLD')}</row>`,
  `<row r="11">${str('A11', 'Chữ đỏ')}${fml('B11', 'DVH.Font.Color(A11,255)', 'OLD')}</row>`,
]
const zip = new JSZip()
zip.file(
  '[Content_Types].xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
)
zip.file(
  '_rels/.rels',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
)
zip.file(
  'xl/workbook.xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>',
)
zip.file(
  'xl/_rels/workbook.xml.rels',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
)
zip.file('xl/styles.xml', styles)
zip.file(
  'xl/worksheets/sheet1.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="9" width="16" customWidth="1"/></cols><sheetData>${rows.join('')}</sheetData></worksheet>`,
)
const xlsxPath = join(outDir, 's5-probe.xlsx')
writeFileSync(xlsxPath, await zip.generateAsync({ type: 'nodebuffer' }))

// ---------------- drive the app ----------------
const profile = join(outDir, `profile-${Date.now()}`)
mkdirSync(profile, { recursive: true })
const env = { ...process.env, GENOFFICE_LANG: 'en', GENOFFICE_DEBUG_HOOKS: '1' }
delete env.ELECTRON_RUN_AS_NODE
const app = await electron.launch({
  executablePath: exe,
  args: [`--user-data-dir=${profile}`],
  env,
  timeout: 60000,
})
const pid = app.process().pid
try {
  const home = await app.firstWindow()
  await home.getByText('Welcome to DVH Office', { exact: true }).waitFor({ timeout: 30000 })
  await home.keyboard.press('Escape')
  await home.evaluate((p) => window.aiOffice.openPath(p), xlsxPath)
  let page
  for (let i = 0; i < 150 && !page; i++) {
    page = app
      .context()
      .pages()
      .find((p) => p.url().includes('://sheets/'))
    if (!page) await sleep(200)
  }
  const errors = []
  page.on('pageerror', (e) => errors.push(e.message))
  await page.getByRole('button', { name: 'Formulas', exact: true }).waitFor({ timeout: 30000 })
  await sleep(4000)

  const values = (a1) =>
    page.evaluate(
      (a) =>
        window.__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveSheet()
          .getRange(a)
          .getValues(),
      a1,
    )
  const styleOf = (cells) =>
    page.evaluate((list) => {
      const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      const pick = (s) =>
        s && {
          bl: s.bl,
          it: s.it,
          ul: s.ul?.s,
          ulType: s.ul?.t,
          bg: s.bg?.rgb,
          cl: s.cl?.rgb,
          n: s.n?.pattern,
        }
      return Object.fromEntries(list.map((a) => [a, pick(ws.getRange(a).getCellStyleData())]))
    }, cells)
  const act = (fn, arg) => page.evaluate(fn, arg)
  const undoDisabled = () =>
    page.getByRole('button', { name: 'Undo', exact: true }).first().isDisabled()

  report.initial = {
    table: await values('F1:I4'),
    statuses: await values('B8:B11'),
    styles: await styleOf(['F1', 'G2', 'F3', 'I3', 'I2', 'A8', 'A9', 'A10', 'A11', 'A12']),
    undoDisabledAfterLoad: await undoDisabled(),
  }
  await page.screenshot({ path: join(outDir, 'app-1-initial.png') })

  // the same DVH.Table typed in the app (a journaled formula, not a file formula)
  await page.evaluate(() =>
    window.__genofficeControl({ cmd: 'goto', target: { kind: 'range', range: 'K1' } }),
  )
  await page.keyboard.type('=DVH.Table(A1:D5,1,,C1:C5="Đạt")')
  await page.keyboard.press('Enter')
  await sleep(2000)
  report.typed = { table: await values('K1:N4'), styles: await styleOf(['K1', 'L2', 'K3', 'N3']) }

  // a source format change flows to the output (G2 mirrors B2)
  await act(() =>
    window.__genofficeDebug.univerAPI
      .getActiveWorkbook()
      .getActiveSheet()
      .getRange('B2')
      .setFontWeight('bold'),
  )
  await sleep(1500)
  report.afterSourceBold = await styleOf(['B2', 'G2'])

  // a source value change re-filters: CV02 now passes and brings its yellow fill
  await act(() =>
    window.__genofficeDebug.univerAPI
      .getActiveWorkbook()
      .getActiveSheet()
      .getRange('C3')
      .setValue('Đạt'),
  )
  await sleep(2000)
  report.afterSourceValue = {
    table: await values('F1:I5'),
    styles: await styleOf(['F3', 'I3', 'F4', 'I5']),
  }

  // deleting the DVH.FillColor formula withdraws its fill
  await act(() =>
    window.__genofficeDebug.univerAPI
      .getActiveWorkbook()
      .getActiveSheet()
      .getRange('B10')
      .setValue(''),
  )
  await sleep(1500)
  report.afterFormulaDeleted = await styleOf(['A10'])
  await page.screenshot({ path: join(outDir, 'app-2-after-edits.png') })

  // undo history holds only the user edits (typed K1 + 3 edits): three undos restore the pre-edit state
  for (let i = 0; i < 3; i++) {
    await page.keyboard.press('Control+z')
    await sleep(700)
  }
  report.afterThreeUndos = {
    undoDisabled: await undoDisabled(),
    table: await values('F1:I5'),
    styles: await styleOf(['A10', 'G2']),
  }
  // redo them back for the save check
  for (let i = 0; i < 3; i++) {
    await page.keyboard.press('Control+y')
    await sleep(700)
  }
  await sleep(1000)
  await page.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(3500)
  report.pageErrors = errors
} finally {
  await app.close().catch(() => {})
  try {
    execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
  } catch {}
}

// ---------------- the saved file ----------------
const saved = await JSZip.loadAsync(readFileSync(xlsxPath))
const sheet = await saved.file('xl/worksheets/sheet1.xml').async('string')
const stylesXml = await saved.file('xl/styles.xml').async('string')
const xfs = [
  ...(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? '').matchAll(
    /<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g,
  ),
].map((m) => m[0])
const fonts = [
  ...(/<fonts[^>]*>([\s\S]*?)<\/fonts>/.exec(stylesXml)?.[1] ?? '').matchAll(
    /<font\b[\s\S]*?<\/font>|<font\b[^>]*\/>/g,
  ),
].map((m) => m[0])
const fills = [
  ...(/<fills[^>]*>([\s\S]*?)<\/fills>/.exec(stylesXml)?.[1] ?? '').matchAll(
    /<fill\b[\s\S]*?<\/fill>/g,
  ),
].map((m) => m[0])
const numFmts = Object.fromEntries(
  [...stylesXml.matchAll(/<numFmt\b[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g)].map((m) => [
    m[1],
    m[2],
  ]),
)
const attr = (xml, name) => new RegExp(`\\b${name}="([^"]*)"`).exec(xml)?.[1]
function cellStyle(ref) {
  const cell = new RegExp(`<c r="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`).exec(sheet)?.[0]
  if (!cell) return null
  const xf = xfs[Number(attr(cell.slice(0, cell.indexOf('>') + 1), 's') ?? 0)] ?? ''
  const font = fonts[Number(attr(xf, 'fontId') ?? 0)] ?? ''
  const fill = fills[Number(attr(xf, 'fillId') ?? 0)] ?? ''
  return {
    bold: /<b\/>|<b val="1"/.test(font),
    italic: /<i\/>|<i val="1"/.test(font),
    fontColor: /<color rgb="([^"]+)"/.exec(font)?.[1],
    fill: /<fgColor rgb="([^"]+)"/.exec(fill)?.[1],
    numFmt: numFmts[attr(xf, 'numFmtId')] ?? attr(xf, 'numFmtId'),
    formula: /<f[^>]*>([^<]*)<\/f>/.exec(cell)?.[1],
    value: /<v>([^<]*)<\/v>/.exec(cell)?.[1] ?? /<t[^>]*>([^<]*)<\/t>/.exec(cell)?.[1],
  }
}
report.savedFile = Object.fromEntries(
  [
    'F1',
    'G1',
    'F2',
    'G2',
    'F3',
    'I3',
    'F4',
    'I4',
    'F5',
    'A8',
    'A9',
    'A10',
    'A11',
    'B8',
    'B10',
    'B11',
    'K1',
    'L1',
    'L2',
    'K3',
    'N3',
    'K5',
    'N5',
  ].map((r) => [r, cellStyle(r)]),
)
writeFileSync(join(outDir, 'app-report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

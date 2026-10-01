// Spike S4: write a 1,000-row × 10-column table (values + per-cell styles) into
// Univer as one undo item, then render the same table through DVH.Table's
// format channel (10k cell formats), and save. Measures time and JS heap.
// Usage: node app-check.mjs <outDir> <path to DVH Office.exe> [rows]
import { _electron as electron } from 'playwright-core'
import JSZip from 'jszip'
import { mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { execFileSync } from 'node:child_process'

const outDir = resolve(process.argv[2])
const exe = resolve(process.argv[3])
const ROWS = Number(process.argv[4] ?? 1000)
mkdirSync(outDir, { recursive: true })
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const report = { rows: ROWS, columns: 10 }

const zip = new JSZip()
zip.file(
  '[Content_Types].xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
)
zip.file(
  '_rels/.rels',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
)
zip.file(
  'xl/workbook.xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
)
zip.file(
  'xl/_rels/workbook.xml.rels',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
)
zip.file(
  'xl/worksheets/sheet1.xml',
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>S4</t></is></c></row></sheetData></worksheet>',
)
const path = join(outDir, 's4-table.xlsx')
writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))

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
  await home.evaluate((p) => window.aiOffice.openPath(p), path)
  let page
  for (let i = 0; i < 150 && !page; i++) {
    page = app
      .context()
      .pages()
      .find((p) => p.url().includes('://sheets/'))
    if (!page) await sleep(200)
  }
  await page.getByRole('button', { name: 'Formulas', exact: true }).waitFor({ timeout: 60000 })
  await page.waitForFunction(() => window.__genofficeDebug?.univerAPI?.getActiveWorkbook(), null, {
    timeout: 60000,
  })
  await sleep(2000)

  // 1. one batch write: values + per-cell styles, timed until two frames have painted
  report.write = await page.evaluate(async (rows) => {
    const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
    // a blank sheet has 1,000 rows; growing it is its own undo item, outside the timing
    if (ws.getMaxRows() < rows + 10)
      ws.insertRowsAfter(ws.getMaxRows() - 1, rows + 10 - ws.getMaxRows())
    await new Promise((r) => setTimeout(r, 500))
    const border = { s: 1, cl: { rgb: '#808080' } }
    const bd = { t: border, b: border, l: border, r: border }
    const header = Array.from({ length: 10 }, (_, c) => ({
      v: `Cột ${c + 1}`,
      s: { bl: 1, bg: { rgb: '#D9E2F3' }, bd },
    }))
    const body = Array.from({ length: rows }, (_, r) =>
      Array.from({ length: 10 }, (_, c) => ({
        v: c < 2 ? `CV${String(r + 1).padStart(4, '0')}-${c}` : (r + 1) * (c + 0.25),
        s: {
          bd,
          ...(r % 2 ? { bg: { rgb: '#F2F2F2' } } : {}),
          ...(c >= 2 ? { n: { pattern: '#,##0.00' } } : {}),
          ...(r % 50 === 0 ? { cl: { rgb: '#C00000' }, bl: 1 } : {}),
        },
      })),
    )
    const frames = () =>
      new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done)))
    const heapBefore = performance.memory?.usedJSHeapSize ?? 0
    const t0 = performance.now()
    ws.getRange(0, 0, rows + 1, 10).setValues([header, ...body])
    const tCommand = performance.now()
    await frames()
    const tPainted = performance.now()
    return {
      commandMs: Math.round(tCommand - t0),
      paintedMs: Math.round(tPainted - t0),
      heapDeltaMB: Math.round(((performance.memory?.usedJSHeapSize ?? 0) - heapBefore) / 1048576),
      lastCell: ws.getRange(rows, 9).getValue(),
      styleB500: ws.getRange(499, 1).getCellStyleData(),
    }
  }, ROWS)

  // 2. one undo reverts the whole batch, one redo brings it back
  await page.keyboard.press('Control+z')
  await sleep(1500)
  const afterUndo = await page.evaluate(
    (rows) =>
      window.__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(rows, 9)
        .getValue(),
    ROWS,
  )
  const t = Date.now()
  await page.keyboard.press('Control+y')
  await page.waitForFunction(
    (rows) =>
      window.__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(rows, 9)
        .getValue() != null,
    ROWS,
    { timeout: 60000 },
  )
  report.undoRedo = { lastCellAfterUndo: afterUndo, redoMs: Date.now() - t }

  // 3. the same table through DVH.Table (spill + 10k formats on the channel)
  report.dvhTable = await page.evaluate(async (rows) => {
    const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
    const t0 = performance.now()
    ws.getRange(0, 11).setFormula(`=DVH.Table(A1:J${rows + 1},1,,)`)
    for (let i = 0; i < 600; i++) {
      if (ws.getRange(rows, 20).getValue() != null) break
      await new Promise((r) => setTimeout(r, 50))
    }
    const tValues = performance.now()
    await new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done)))
    return {
      spilledMs: Math.round(tValues - t0),
      paintedMs: Math.round(performance.now() - t0),
      lastSpilled: ws.getRange(rows, 20).getValue(),
      styleM500: ws.getRange(499, 12).getCellStyleData(),
    }
  }, ROWS)

  // scroll through the formatted spill: average frame time while the viewport moves
  report.scroll = await page.evaluate(async () => {
    const times = []
    let last = performance.now()
    for (let i = 0; i < 40; i++) {
      window.dispatchEvent(new Event('resize'))
      await new Promise((r) => requestAnimationFrame(r))
      const now = performance.now()
      times.push(now - last)
      last = now
    }
    return { avgFrameMs: Math.round(times.reduce((a, b) => a + b, 0) / times.length) }
  })
  await page.screenshot({ path: join(outDir, 's4-table.png') })

  // 4. save (journal of the batch + 10k baked format edits)
  const t1 = Date.now()
  await page.getByRole('button', { name: /^Save \(/ }).click()
  await page.waitForFunction(
    () =>
      [...document.querySelectorAll('.app-toast')].some((el) =>
        /Saved|failed|error/i.test(el.textContent ?? ''),
      ),
    null,
    { timeout: 180000 },
  )
  report.save = {
    ms: Date.now() - t1,
    toasts: await page.evaluate(() =>
      [...document.querySelectorAll('.app-toast')].map((el) => el.textContent?.trim()),
    ),
    bytes: statSync(path).size,
  }
} finally {
  await app.close().catch(() => {})
  try {
    execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
  } catch {}
}
const saved = await JSZip.loadAsync(readFileSync(path))
const sheet = await saved.file('xl/worksheets/sheet1.xml').async('string')
report.savedFile = {
  cells: (sheet.match(/<c r=/g) ?? []).length,
  styledCells: (sheet.match(/<c r="[A-Z]+\d+" s="[1-9]/g) ?? []).length,
  cellXfs: Number(
    /<cellXfs count="(\d+)"/.exec(await saved.file('xl/styles.xml').async('string'))?.[1] ?? 0,
  ),
}
writeFileSync(join(outDir, 'app-report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

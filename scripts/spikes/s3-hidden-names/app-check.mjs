// Spike S3: hidden DVH names (_dvh.f.*, _dvh.t.*) under real editor operations in
// the packaged app — row insert, cut/paste, sheet rename, deleting the bound row,
// and a row insert while a large workbook is still streaming.
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

const FIELD = '_dvh.f.f_projectname0001'
const TABLE = '_dvh.t.t_workitems000001'

async function buildWorkbook(path, extraRows) {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/></sheets>' +
      `<definedNames><definedName name="${FIELD}" hidden="1">Data!$B$5</definedName><definedName name="${TABLE}" hidden="1">Data!$A$8:$C$11</definedName><definedName name="TenDuAn">Data!$B$5</definedName></definedNames></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
  )
  const rows = [
    '<row r="5"><c r="A5" t="inlineStr"><is><t>Tên dự án</t></is></c><c r="B5" t="inlineStr"><is><t>Dự án A</t></is></c></row>',
    '<row r="8"><c r="A8" t="inlineStr"><is><t>Mã</t></is></c><c r="B8" t="inlineStr"><is><t>Công việc</t></is></c><c r="C8" t="inlineStr"><is><t>KL</t></is></c></row>',
  ]
  for (let r = 20; r < 20 + extraRows; r++) {
    rows.push(
      `<row r="${r}"><c r="A${r}"><v>${r}</v></c><c r="B${r}"><v>${r * 3}</v></c><c r="C${r}"><v>${r % 97}</v></c><c r="D${r}"><v>${r * 0.5}</v></c><c r="E${r}"><v>${r % 13}</v></c></row>`,
    )
  }
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join('')}</sheetData></worksheet>`,
  )
  zip.file(
    'xl/worksheets/sheet2.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

async function namesIn(path) {
  const zip = await JSZip.loadAsync(readFileSync(path))
  const xml = await zip.file('xl/workbook.xml').async('string')
  return Object.fromEntries(
    [...xml.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)].map((m) => [
      /name="([^"]*)"/.exec(m[1])[1],
      m[2] + (/hidden="1"/.test(m[1]) ? ' (hidden)' : ''),
    ]),
  )
}

async function withSheets(path, fn) {
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
    // the debug hook is attached once the first viewport of the file is installed
    await page.waitForFunction(
      () => window.__genofficeDebug?.univerAPI?.getActiveWorkbook(),
      null,
      { timeout: 60000 },
    )
    const ws = (code) =>
      page.evaluate(
        `(() => { const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet(); ${code} })()`,
      )
    const save = async () => {
      await page.getByRole('button', { name: /^Save \(/ }).click()
      await sleep(3500)
      return page.evaluate(() =>
        [...document.querySelectorAll('.app-toast')]
          .map((t) => t.textContent?.trim())
          .filter(Boolean),
      )
    }
    return await fn({ page, ws, save })
  } finally {
    await app.close().catch(() => {})
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
    } catch {}
  }
}

// ---------------- small workbook: one operation per save ----------------
const small = join(outDir, 's3-small.xlsx')
await buildWorkbook(small, 0)
await withSheets(small, async ({ page, ws, save }) => {
  await sleep(3000)
  await ws('ws.insertRowBefore(2)')
  await sleep(800)
  report.insertRow = { toasts: await save(), names: await namesIn(small) }

  // cut B6 (the bound cell after the insert) and paste it to E6
  await page.evaluate(() =>
    window.__genofficeControl({ cmd: 'goto', target: { kind: 'range', range: 'B6' } }),
  )
  await page.keyboard.press('Control+x')
  await sleep(500)
  await page.evaluate(() =>
    window.__genofficeControl({ cmd: 'goto', target: { kind: 'range', range: 'E6' } }),
  )
  await page.keyboard.press('Control+v')
  await sleep(1000)
  report.cutPaste = {
    e6: await ws("return ws.getRange('E6').getValue()"),
    b6: await ws("return ws.getRange('B6').getValue()"),
    toasts: await save(),
    names: await namesIn(small),
  }

  await ws("ws.setName('Dữ liệu')")
  await sleep(800)
  report.renameSheet = { toasts: await save(), names: await namesIn(small) }

  // delete the row that holds the original binding (row 6 after the insert)
  const before = readFileSync(small)
  await ws('ws.deleteRow(5)')
  await sleep(800)
  const toasts = await save()
  report.deleteBoundRow = {
    toasts,
    fileUnchanged: Buffer.compare(before, readFileSync(small)) === 0,
    names: await namesIn(small),
  }
  await page.screenshot({ path: join(outDir, 'delete-bound-row.png') })
})
writeFileSync(join(outDir, 'app-report.json'), JSON.stringify(report, null, 2))

// ---------------- large workbook: insert while it still streams ----------------
const large = join(outDir, 's3-large.xlsx')
await buildWorkbook(large, 150_000)
await withSheets(large, async ({ ws, save }) => {
  await sleep(500)
  const loading = await ws('return ws.getMaxRows()')
  await ws('ws.insertRowBefore(0)')
  await sleep(800)
  report.largeInsertWhileLoading = {
    rowsKnownAtInsert: loading,
    toasts: await save(),
    names: await namesIn(large),
  }
})

writeFileSync(join(outDir, 'app-report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

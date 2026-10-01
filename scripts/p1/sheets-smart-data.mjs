// P1 acceptance (Sheets): bind a cell to a Smart Data field through the panel,
// move it (row insert, cut/paste), edit it, save, delete the bound row, save
// again, then reopen. Usage: node sheets-smart-data.mjs <outDir> <DVH Office.exe>
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
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="5"><c r="A5" t="inlineStr"><is><t>Tên dự án</t></is></c><c r="B5" t="inlineStr"><is><t>Dự án A</t></is></c></row></sheetData></worksheet>',
)
const path = join(outDir, 'p1-sheets.xlsx')
writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))

async function fileState() {
  const z = await JSZip.loadAsync(readFileSync(path))
  const wb = await z.file('xl/workbook.xml').async('string')
  const parts = {}
  for (const p of Object.keys(z.files).filter((p) => /^customXml\/item\d+\.xml$/.test(p))) {
    parts[p] = await z.file(p).async('string')
  }
  const model = Object.values(parts).find((x) => x.includes('urn:dvh-office:model:1')) ?? null
  const history = Object.values(parts).find((x) => x.includes('urn:dvh-office:history:1')) ?? null
  return {
    names: [...wb.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)].map(
      (m) =>
        `${/name="([^"]*)"/.exec(m[1])[1]}=${m[2]}${/hidden="1"/.test(m[1]) ? ' (hidden)' : ''}`,
    ),
    modelField: model
      ? (/<dvh:f\b[^>]*name="Project\.Name"[^>]*>([^<]*)</.exec(model)?.[1] ?? null)
      : null,
    historyActions: history ? [...history.matchAll(/"action":"([^"]+)"/g)].map((m) => m[1]) : [],
  }
}

async function withApp(fn) {
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
    await page.waitForFunction(
      () => window.__genofficeDebug?.univerAPI?.getActiveWorkbook(),
      null,
      { timeout: 60000 },
    )
    await sleep(2500)
    const errors = []
    page.on('pageerror', (e) => errors.push(e.message))
    const ws = (code) =>
      page.evaluate(
        `(() => { const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet(); ${code} })()`,
      )
    const names = () =>
      page.evaluate(() =>
        window.__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getDefinedNames()
          .map((n) => `${n.getName()}=${n.getFormulaOrRefString()}`),
      )
    const goto = (range) =>
      page.evaluate(
        (r) => window.__genofficeControl({ cmd: 'goto', target: { kind: 'range', range: r } }),
        range,
      )
    const toasts = () =>
      page.evaluate(() =>
        [...document.querySelectorAll('.app-toast')]
          .map((t) => t.textContent?.trim())
          .filter(Boolean),
      )
    const status = () =>
      page.evaluate(() => document.querySelector('.status-msg')?.textContent?.trim() ?? '')
    const save = async () => {
      await page.getByRole('button', { name: /^Save \(/ }).click()
      await sleep(3500)
      await page.waitForFunction(
        () => window.__genofficeDebug?.univerAPI?.getActiveWorkbook(),
        null,
        { timeout: 60000 },
      )
      await sleep(1500)
      return status()
    }
    // keyboard input belongs to the grid, not to a panel input that kept focus
    const focusGrid = async () => {
      const box = await page.evaluate(() => {
        const canvases = [...document.querySelectorAll('canvas')].map((c) =>
          c.getBoundingClientRect(),
        )
        const big = canvases.sort((a, b) => b.width * b.height - a.width * a.height)[0]
        return big ? { x: big.left + 120, y: big.top + 300 } : null
      })
      if (box) await page.mouse.click(box.x, box.y)
      await sleep(200)
    }
    return await fn({ page, ws, names, goto, toasts, save, errors, focusGrid, status })
  } finally {
    await app.close().catch(() => {})
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
    } catch {}
  }
}

await withApp(async ({ page, ws, names, goto, toasts, save, errors, focusGrid }) => {
  await page.getByRole('button', { name: 'Smart Data', exact: true }).click()
  const panel = page.getByRole('dialog', { name: 'Smart Data fields' })
  await panel.waitFor()
  report.emptyPanel = await panel.innerText()
  await goto('B5')
  await panel.getByRole('textbox').fill('Project.Name')
  await panel.getByRole('button', { name: 'Bind selected cell' }).click()
  await sleep(800)
  report.afterBind = {
    panel: await panel.innerText(),
    names: await names(),
    toasts: await toasts(),
  }

  await ws('ws.insertRowBefore(2)')
  await sleep(800)
  report.afterInsertRow = await names()

  await focusGrid()
  await goto('B6')
  await page.keyboard.press('Control+x')
  await sleep(400)
  await goto('E6')
  await page.keyboard.press('Control+v')
  await sleep(1000)
  report.afterCutPaste = await names()

  await goto('E6')
  await page.keyboard.type('Dự án B')
  await page.keyboard.press('Enter')
  await sleep(1200)
  report.panelAfterEdit = await panel.innerText()
  await page.screenshot({ path: join(outDir, 'panel.png') })
  report.saveStatus1 = await save()
  report.saved1 = await fileState()

  await ws('ws.deleteRow(5)')
  await sleep(1200)
  report.afterDeleteBoundRow = { toasts: await toasts(), names: await names() }
  const saveStatus2 = await save()
  report.saved2 = { status: saveStatus2, toasts: await toasts(), file: await fileState() }
  report.pageErrors = errors
})

// reopen: the field and its (broken) binding come back from the file
await withApp(async ({ page }) => {
  await page.getByRole('button', { name: 'Smart Data', exact: true }).click()
  const panel = page.getByRole('dialog', { name: 'Smart Data fields' })
  await panel.waitFor()
  await sleep(1500)
  report.reopenedPanel = await panel.innerText()
  await page.getByRole('button', { name: 'Formulas', exact: true }).click()
})

writeFileSync(join(outDir, 'report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

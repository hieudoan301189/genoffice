// Spike S1, last leg: the Office-saved files reopened and re-saved by the packaged
// DVH Office app (Univer + ProseMirror front ends, not just the engines).
// Usage: node app-check.mjs <outDir> <path to DVH Office.exe>
import { _electron as electron } from 'playwright-core'
import JSZip from 'jszip'
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { execFileSync } from 'node:child_process'

const outDir = resolve(process.argv[2])
const exe = resolve(process.argv[3])
const report = {}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

async function withApp(fn) {
  const profile = join(outDir, `app-profile-${Date.now()}`)
  mkdirSync(profile, { recursive: true })
  const env = { ...process.env, GENOFFICE_LANG: 'en' }
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
    return await fn(app, home)
  } finally {
    await app.close().catch(() => {})
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
    } catch {}
  }
}

async function openIn(app, home, path, scheme) {
  await home.evaluate((p) => window.aiOffice.openPath(p), path)
  for (let i = 0; i < 150; i++) {
    const page = app
      .context()
      .pages()
      .find((p) => p.url().includes(`://${scheme}/`))
    if (page) return page
    await sleep(200)
  }
  throw new Error(`no ${scheme} page`)
}

async function control(page, req) {
  for (let i = 0; i < 100; i++) {
    const reply = await page.evaluate(
      (r) => window.__genofficeControl?.(r) ?? { status: 'not_ready' },
      req,
    )
    if (reply.status !== 'not_ready') return reply
    await sleep(200)
  }
  throw new Error('control not ready')
}

async function readRange(page, range) {
  await control(page, { cmd: 'goto', target: { kind: 'range', range } })
  return (await control(page, { cmd: 'selection' })).result
}

// ---------------- xlsx ----------------
const xlsxPath = join(outDir, 'app-excel-saved-dvh-s1.xlsx')
copyFileSync(join(outDir, 'excel-saved-dvh-s1.xlsx'), xlsxPath)
await withApp(async (app, home) => {
  const page = await openIn(app, home, xlsxPath, 'sheets')
  page.on('pageerror', (e) => console.log('PAGE ERROR', e.message))
  await page.getByRole('button', { name: 'Formulas', exact: true }).waitFor({ timeout: 30000 })
  await sleep(4000)
  // after Excel inserted row 3: B6 = project, C6 = =_dvh.f..., D6 = =TenDuAn
  report.xlsxRow6 = await readRange(page, 'B6:D6')
  // does C6 (=hidden name) recalc, or only show Excel's cached value? D6 (=visible name) is the control
  // E6 = plain reference control, F6 = the hidden name typed fresh in DVH Office
  for (const [ref, formula] of [
    ['E6', '=B6'],
    ['F6', '=_dvh.f.f_projectname0001'],
    ['G6', '=TenDuAn'],
  ]) {
    await control(page, { cmd: 'goto', target: { kind: 'range', range: ref } })
    await page.keyboard.type(formula)
    await page.keyboard.press('Enter')
  }
  await sleep(1500)
  report.xlsxRow6WithControls = await readRange(page, 'B6:G6')
  await control(page, { cmd: 'goto', target: { kind: 'range', range: 'B6' } })
  await page.keyboard.type('Changed in DVH Office')
  await page.keyboard.press('Enter')
  await sleep(4000)
  report.xlsxRow6AfterB6Edit = await readRange(page, 'B6:G6')
  await control(page, { cmd: 'goto', target: { kind: 'range', range: 'A1' } })
  await page.keyboard.type('Edited in DVH Office')
  await page.keyboard.press('Enter')
  await sleep(500)
  report.xlsxA1 = await readRange(page, 'A1')
  await page.screenshot({ path: join(outDir, 'app-xlsx.png') })
  await page.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(3000)
})
{
  const zip = await JSZip.loadAsync(readFileSync(xlsxPath))
  const workbook = await zip.file('xl/workbook.xml').async('string')
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string')
  report.xlsxSaved = {
    hiddenNames: [...workbook.matchAll(/<definedName\b[^>]*>[^<]*<\/definedName>/g)].map(
      (m) => m[0],
    ),
    customXml: Object.keys(zip.files).filter((p) => p.startsWith('customXml/')),
    a1: /<c r="A1"[\s\S]*?<\/c>/.exec(sheet)?.[0],
    c6: /<c r="C6"[\s\S]*?<\/c>/.exec(sheet)?.[0],
  }
}

// ---------------- docx ----------------
const docxPath = join(outDir, 'app-word-saved-dvh-s1.docx')
copyFileSync(join(outDir, 'word-saved-dvh-s1.docx'), docxPath)
await withApp(async (app, home) => {
  const page = await openIn(app, home, docxPath, 'docs')
  page.on('pageerror', (e) => console.log('PAGE ERROR', e.message))
  const editor = page.locator('.ProseMirror').first()
  await editor.waitFor({ timeout: 30000 })
  await sleep(3000)
  report.docxText = (await editor.innerText()).slice(0, 600)
  await editor.click()
  await page.keyboard.press('Control+End')
  await page.keyboard.type(' (DVH Office)')
  await sleep(500)
  await page.screenshot({ path: join(outDir, 'app-docx.png') })
  await page.keyboard.press('Control+s')
  await sleep(3000)
})
{
  const zip = await JSZip.loadAsync(readFileSync(docxPath))
  const doc = await zip.file('word/document.xml').async('string')
  report.docxSaved = {
    customXml: Object.keys(zip.files).filter((p) => p.startsWith('customXml/')),
    fieldSdt: doc.includes('w:tag w:val="dvh:f:f_projectname0001"'),
    dataBinding: doc.includes('w:storeItemID="{6F1E2C1A-4B7D-4E39-9C1A-5D2B7E8F9A01}"'),
    tableSdt: doc.includes('w:tag w:val="dvh:t:t_workitems000001"'),
    typed: doc.includes('(DVH Office)'),
  }
}

writeFileSync(join(outDir, 'app-report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

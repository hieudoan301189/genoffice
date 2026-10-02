// P3 acceptance: automatic links. One app session holds the workbook (Sheets)
// and the report (Docs). The report links the workbook in Automatic mode, then
// 1. edits in Sheets (values, a new row, the title fill, a field) reach the
//    report live, without saving the workbook;
// 2. the workbook is edited and saved by real Excel: the watched file updates
//    the report (read from the cells, Excel keeps the DVH model part stale).
// Usage: node auto-link.mjs <outDir> <DVH Office.exe>
import { _electron as electron } from 'playwright-core'
import JSZip from 'jszip'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const outDir = resolve(process.argv[2])
const exe = resolve(process.argv[3])
mkdirSync(outDir, { recursive: true })
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const report = {}
const step = (label) => console.error(`${new Date().toISOString()} ${label}`)
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

async function writeXlsx(path) {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>`,
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function writeDocx(path, paragraphs) {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  const ps = paragraphs
    .map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`)
    .join('')
  zip.file(
    'word/document.xml',
    `${XML}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${ps}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

const xlsx = join(outDir, 'p3-khoi-luong.xlsx')
const docx = join(outDir, 'p3-bao-cao.docx')
await writeXlsx(xlsx)
await writeDocx(docx, ['Báo cáo khối lượng', 'Dự án: ', 'Hết.'])

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
const findPage = async (kind) => {
  for (let i = 0; i < 150; i++) {
    const page = app
      .context()
      .pages()
      .find((p) => p.url().includes(`://${kind}/`))
    if (page) return page
    await sleep(200)
  }
  throw new Error(`no ${kind} page`)
}

try {
  const home = await app.firstWindow()
  await home.getByText('Welcome to DVH Office', { exact: true }).waitFor({ timeout: 30000 })
  await home.keyboard.press('Escape')

  // ---- the workbook: a styled range becomes WorkItems, a field cell, saved once
  step('sheets: open')
  await home.evaluate((p) => window.aiOffice.openPath(p), xlsx)
  const sheets = await findPage('sheets')
  const errors = []
  sheets.on('pageerror', (e) => errors.push(`sheets: ${e.message}`))
  sheets.on('dialog', (d) => d.accept())
  await sheets.waitForFunction(() => window.__genofficeDebug?.dvhActions, null, { timeout: 60000 })
  await sleep(2500)
  const ws = (code) =>
    sheets.evaluate(
      `(() => { const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet(); ${code} })()`,
    )
  const sheetsRun = (name, input) =>
    sheets.evaluate(
      async ({ name, input }) => {
        const r = await window.__genofficeDebug.dvhActions.run(name, input, {
          caller: 'ui',
          docId: 'driver',
          permissions: new Set(['read', 'write']),
        })
        return JSON.parse(JSON.stringify(r.output ?? null))
      },
      { name, input },
    )
  await ws(
    `const head = { bl: 1, bg: { rgb: '#FFF2CC' }, bd: { b: { s: 1, cl: { rgb: '#808080' } } } }
     const qty = { cl: { rgb: '#C00000' }, n: { pattern: '#,##0.00' } }
     ws.getRange('A1:C4').setValues([
       [{ v: 'Mã', s: head }, { v: 'Công việc', s: head }, { v: 'Khối lượng', s: head }],
       [{ v: 'AB.1' }, { v: 'Đào móng' }, { v: 120.5, s: qty }],
       [{ v: 'AB.2' }, { v: 'Bê tông lót' }, { v: 18, s: qty }],
       [{ v: 'AB.3' }, { v: 'Cốt thép' }, { v: 2450.25, s: qty }],
     ])
     ws.getRange('F1').setValue('Dự án A')`,
  )
  await sleep(500)
  await sheetsRun('Data.CreateCollection', { name: 'WorkItems', range: 'A1:C4' })
  await sheetsRun('Spreadsheet.BindField', { name: 'Project.Name', cell: 'F1' })
  await sheets.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(4000)

  // ---- the report: link (automatic), a field and a table
  step('docs: open and link')
  await home.evaluate((p) => window.aiOffice.openPath(p), docx)
  const docs = await findPage('docs')
  docs.on('pageerror', (e) => errors.push(`docs: ${e.message}`))
  docs.on('dialog', (d) => d.accept())
  await docs.locator('.ProseMirror').first().waitFor({ timeout: 30000 })
  await docs.waitForFunction(() => window.__dvhActions, null, { timeout: 30000 })
  await sleep(2000)
  const docsRun = (name, input) =>
    docs.evaluate(
      async ({ name, input }) => {
        try {
          const r = await window.__dvhActions.run(name, input, {
            caller: 'ui',
            docId: 'driver',
            permissions: new Set(['read', 'write']),
          })
          return JSON.parse(JSON.stringify(r.output ?? null))
        } catch (e) {
          return { error: e.message }
        }
      },
      { name, input },
    )
  await docs.getByText('Insert', { exact: true }).first().click()
  await docs.locator('button', { hasText: 'Smart Data' }).first().click()
  const panel = docs.getByRole('dialog', { name: 'Smart Data fields' })
  await panel.waitFor()
  await app.evaluate(({ dialog }, p) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [p] })
  }, xlsx)
  await panel.getByRole('button', { name: 'Link workbook…' }).click()
  await sleep(1500)
  report.insertField = await docsRun('Document.InsertField', {
    field: 'Project.Name',
    blockIndex: 1,
  })
  report.insertTable = await docsRun('Document.InsertTable', {
    collection: 'WorkItems',
    afterBlockIndex: 1,
  })
  await panel.getByRole('combobox', { name: 'Update' }).selectOption('auto')
  await sleep(1500)

  const docState = () =>
    docs.evaluate(() => ({
      field: document.querySelector('[data-dvh-field]')?.textContent ?? null,
      table: [...(document.querySelector('table[data-dvh-table]')?.rows ?? [])].map((r) =>
        [...r.cells].map((c) => c.textContent.trim()),
      ),
      headerFill: (() => {
        const cell = document.querySelector(
          'table[data-dvh-table] tr td, table[data-dvh-table] tr th',
        )
        return cell ? getComputedStyle(cell).backgroundColor : null
      })(),
      status: [...document.querySelectorAll('[data-status]')].map((e) => e.dataset.status),
    }))
  report.before = await docState()

  // ---- 1. live: edit in Sheets, do not save
  step('live edits in sheets')
  await ws(
    `ws.getRange('C2').setValue(130)
     ws.getRange('A5:C5').setValues([['AB.4', 'Ván khuôn', 75]])
     ws.getRange('A1:C1').setBackground('#C6EFCE')
     ws.getRange('F1').setValue('Dự án B')`,
  )
  await sleep(2500)
  report.live = await docState()
  report.sheetsUnsaved = await sheets.evaluate(
    () => document.querySelector('button[aria-label^="Save ("]')?.disabled === false,
  )

  // ---- 2. real Excel edits the saved workbook
  step('save sheets, then edit with Excel')
  await sheets.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(4000)
  const editsFile = join(outDir, 'excel-edits.json')
  writeFileSync(
    editsFile,
    JSON.stringify([
      { cell: 'B3', value: 'Bê tông lót móng' },
      { cell: 'A6', value: 'AB.5' },
      { cell: 'B6', value: 'Xây gạch' },
      { cell: 'C6', value: 32.4 },
      { range: 'A1:C1', fill: '#BDD7EE' },
    ]),
  )
  report.excel = execFileSync(
    'powershell',
    [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      join(here, 'excel-edit.ps1'),
      '-Path',
      xlsx,
      '-EditsFile',
      editsFile,
    ],
    { encoding: 'utf8' },
  ).trim()
  await sleep(4000)
  report.afterExcel = await docState()

  // save the report and read it back in Word
  step('save docs')
  await docs
    .locator('.ProseMirror')
    .first()
    .click({ position: { x: 5, y: 5 } })
  await docs.keyboard.press('Control+s')
  await sleep(4000)
  await docs.screenshot({ path: join(outDir, 'docs-auto.png') })
  report.pageErrors = errors
} finally {
  const closed = await Promise.race([
    app.close().then(
      () => true,
      () => true,
    ),
    sleep(20000).then(() => false),
  ])
  if (!closed) report.closeTimedOut = true
  try {
    execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
  } catch {}
}

writeFileSync(join(outDir, 'report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

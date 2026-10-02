// P3 demo: records (Bandicam) a workbook in DVH Sheets feeding a report in
// DVH Docs through an automatic link — values, a new row, title and column
// formatting, a field and a removed row show up in the report as they are typed,
// without saving the workbook.
// Usage: node demo-record.mjs <outDir> <DVH Office.exe> [bdcam.exe]
import { _electron as electron } from 'playwright-core'
import JSZip from 'jszip'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { execFileSync, spawn } from 'node:child_process'

const outDir = resolve(process.argv[2])
const exe = resolve(process.argv[3])
const bandicam = process.argv[4] ?? 'C:\\Program Files (x86)\\Bandicam\\bdcam.exe'
mkdirSync(outDir, { recursive: true })
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
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
    `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Khối lượng" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="2" max="2" width="22" customWidth="1"/><col min="3" max="3" width="14" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>`,
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function writeDocx(path) {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  const p = (t, bold = false) =>
    `<w:p><w:r>${bold ? '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>' : ''}<w:t xml:space="preserve">${t}</w:t></w:r></w:p>`
  zip.file(
    'word/document.xml',
    `${XML}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${p('BÁO CÁO KHỐI LƯỢNG NGHIỆM THU', true)}${p('Dự án: ')}${p('Bảng khối lượng (liên kết tự động từ bảng tính):')}${p('Hết.')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

const xlsx = join(outDir, 'demo-khoi-luong.xlsx')
const docx = join(outDir, 'demo-bao-cao.docx')
await writeXlsx(xlsx)
await writeDocx(docx)

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
let recording = false

// Bandicam records its last rectangle: point it at the whole primary screen for
// the demo and give the user's own rectangle back afterwards
const BANDICAM_KEY = 'HKCU:\\Software\\BANDISOFT\\BANDICAM\\OPTION'
const RECT = ['TargetRect.left', 'TargetRect.top', 'TargetRect.right', 'TargetRect.bottom']
const ps = (command) =>
  execFileSync('powershell', ['-NoProfile', '-Command', command], { encoding: 'utf8' }).trim()
const savedRect = ps(
  `$p = Get-ItemProperty '${BANDICAM_KEY}'; @(${RECT.map((n) => `$p.'${n}'`).join(',')}) -join ','`,
)
writeFileSync(join(outDir, 'bandicam-rect-backup.txt'), savedRect)
const physical = ps(
  'Add-Type -TypeDefinition \'using System.Runtime.InteropServices; public static class D { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("user32.dll")] public static extern int GetSystemMetrics(int n); }\'; [void][D]::SetProcessDPIAware(); "$([D]::GetSystemMetrics(0)),$([D]::GetSystemMetrics(1))"',
).split(',')
const setRect = (values) =>
  ps(
    RECT.map(
      (n, i) => `Set-ItemProperty '${BANDICAM_KEY}' -Name '${n}' -Value ${values[i]} -Type DWord`,
    ).join('; '),
  )

try {
  const home = await app.firstWindow()
  await home.getByText('Welcome to DVH Office', { exact: true }).waitFor({ timeout: 30000 })
  await home.keyboard.press('Escape')

  step('prepare workbook')
  await home.evaluate((p) => window.aiOffice.openPath(p), xlsx)
  const sheets = await findPage('sheets')
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
          docId: 'demo',
          permissions: new Set(['read', 'write']),
        })
        return JSON.parse(JSON.stringify(r.output ?? null))
      },
      { name, input },
    )
  await ws(
    `const b = { s: 1, cl: { rgb: '#808080' } }
     const bd = { t: b, b: b, l: b, r: b }
     const head = { bl: 1, bg: { rgb: '#FFF2CC' }, bd, ht: 2 }
     const text = { bd }
     const qty = { cl: { rgb: '#C00000' }, n: { pattern: '#,##0.00' }, bd }
     ws.getRange('A1:C5').setValues([
       [{ v: 'Mã hiệu', s: head }, { v: 'Công việc', s: head }, { v: 'Khối lượng', s: head }],
       [{ v: 'AB.11', s: text }, { v: 'Đào móng', s: text }, { v: 120.5, s: qty }],
       [{ v: 'AB.12', s: text }, { v: 'Bê tông lót', s: text }, { v: 18, s: qty }],
       [{ v: 'AF.11', s: text }, { v: 'Cốt thép móng', s: text }, { v: 2450.25, s: qty }],
       [{ v: 'AF.12', s: text }, { v: 'Ván khuôn móng', s: text }, { v: 64.8, s: qty }],
     ])
     ws.getRange('E1').setValue('Tên dự án')
     ws.getRange('F1').setValue('Nhà văn hoá xã Vĩnh Hằng')`,
  )
  await sleep(500)
  await sheetsRun('Data.CreateCollection', { name: 'KhoiLuong', range: 'A1:C5' })
  await sheetsRun('Spreadsheet.BindField', { name: 'Project.Name', cell: 'F1' })
  await sheets.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(4000)

  step('prepare report')
  await home.evaluate((p) => window.aiOffice.openPath(p), docx)
  const docs = await findPage('docs')
  docs.on('dialog', (d) => d.accept())
  await docs.locator('.ProseMirror').first().waitFor({ timeout: 30000 })
  await docs.waitForFunction(() => window.__dvhActions, null, { timeout: 30000 })
  await sleep(2000)
  const docsRun = (name, input) =>
    docs.evaluate(
      async ({ name, input }) => {
        const r = await window.__dvhActions.run(name, input, {
          caller: 'ui',
          docId: 'demo',
          permissions: new Set(['read', 'write']),
        })
        return JSON.parse(JSON.stringify(r.output ?? null))
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
  await docsRun('Document.InsertField', { field: 'Project.Name', blockIndex: 1 })
  await docsRun('Document.InsertTable', { collection: 'KhoiLuong', afterBlockIndex: 2 })
  await panel.getByRole('combobox', { name: 'Update' }).selectOption('auto')
  await sleep(1000)
  await docs
    .locator('.ProseMirror')
    .first()
    .click({ position: { x: 5, y: 5 } })
  await docs.keyboard.press('Control+s')
  await sleep(3000)
  // the panel stays open: it shows the link in Automatic mode

  step('arrange windows')
  await home.evaluate(async () => {
    const tabs = await window.aiOfficeTabs.list()
    const doc = tabs.find((t) => t.kind === 'docs')
    if (doc) await window.aiOfficeTabs.detach(doc.id)
  })
  await sleep(2500)
  await app.evaluate(({ BrowserWindow, screen }) => {
    const area = screen.getPrimaryDisplay().workArea
    const wins = BrowserWindow.getAllWindows().filter((w) => w.isVisible())
    const docsWin = wins.find((w) => /bao-cao|báo cáo/i.test(w.getTitle()))
    const shellWin = wins.find((w) => w !== docsWin)
    const half = Math.floor(area.width / 2)
    shellWin?.setBounds({ x: area.x, y: area.y, width: half, height: area.height })
    docsWin?.setBounds({
      x: area.x + half,
      y: area.y,
      width: area.width - half,
      height: area.height,
    })
    for (const w of [shellWin, docsWin]) {
      w?.setAlwaysOnTop(true, 'screen-saver')
      w?.show()
    }
  })
  await sleep(1500)
  // the report shows its title, the field and the table
  await docs.evaluate(() =>
    document.querySelector('table[data-dvh-table]')?.scrollIntoView({ block: 'center' }),
  )

  // grid focus for typing
  const focusGrid = async () => {
    const box = await sheets.evaluate(() => {
      const big = [...document.querySelectorAll('canvas')]
        .map((c) => c.getBoundingClientRect())
        .sort((a, b) => b.width * b.height - a.width * a.height)[0]
      return big ? { x: big.left + 60, y: big.top + 260 } : null
    })
    if (box) await sheets.mouse.click(box.x, box.y)
  }
  const goto = (range) =>
    sheets.evaluate(
      (r) => window.__genofficeControl({ cmd: 'goto', target: { kind: 'range', range: r } }),
      range,
    )
  const typeSlow = async (text) => {
    await sheets.keyboard.type(text, { delay: 90 })
  }

  step('start recording')
  setRect([0, 0, Number(physical[0]), Number(physical[1])])
  spawn(bandicam, ['/record'], { detached: true, stdio: 'ignore' }).unref()
  recording = true
  await sleep(5000)

  await focusGrid()
  // 1. a quantity changes
  step('demo: quantity')
  await goto('C2')
  await sleep(800)
  await typeSlow('135.75')
  await sheets.keyboard.press('Enter')
  await sleep(3500)

  // 2. a new work item typed under the range
  step('demo: new row')
  await goto('A6')
  await sleep(600)
  await typeSlow('AK.21')
  await sheets.keyboard.press('Tab')
  await typeSlow('Xây tường gạch')
  await sheets.keyboard.press('Tab')
  await typeSlow('32.4')
  await sheets.keyboard.press('Enter')
  await sleep(3500)

  // 3. the title fill and the quantity colour (formatting follows too)
  step('demo: formatting')
  await ws(`ws.getRange('A1:C1').setBackground('#C6EFCE')`)
  await sleep(3000)
  await ws(`ws.getRange('C2:C6').setFontColor('#1F4E79')`)
  await sleep(3500)

  // 4. the project name (a single field)
  step('demo: field')
  await goto('F1')
  await sleep(600)
  await typeSlow('Trường mầm non Vĩnh Hằng')
  await sheets.keyboard.press('Enter')
  await sleep(3500)

  // 5. a work item removed
  step('demo: remove row')
  // Dimension.ROWS: shift the cells below up
  await ws(`ws.getRange('A3:C3').deleteCells(1)`)
  await sleep(4500)

  step('stop recording')
  spawn(bandicam, ['/stop'], { detached: true, stdio: 'ignore' }).unref()
  recording = false
  await sleep(6000)
  await docs.screenshot({ path: join(outDir, 'demo-docs-final.png') })
  await sheets.screenshot({ path: join(outDir, 'demo-sheets-final.png') })
  const final = await docs.evaluate(() => ({
    field: document.querySelector('[data-dvh-field]')?.textContent ?? null,
    table: [...(document.querySelector('table[data-dvh-table]')?.rows ?? [])].map((r) =>
      [...r.cells].map((c) => c.textContent.trim()),
    ),
  }))
  writeFileSync(join(outDir, 'demo-final.json'), JSON.stringify(final, null, 2))
  console.log(JSON.stringify(final, null, 2))
} finally {
  if (recording) spawn(bandicam, ['/stop'], { detached: true, stdio: 'ignore' }).unref()
  await sleep(3000)
  // Bandicam saves its settings on exit: close it first, then restore the user's rectangle
  spawn(bandicam, ['/shutdown'], { detached: true, stdio: 'ignore' }).unref()
  await sleep(5000)
  setRect(savedRect.split(','))
  await Promise.race([app.close().catch(() => {}), sleep(15000)])
  try {
    execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
  } catch {}
}

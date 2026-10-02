// P1 acceptance (Docs ← Sheets): two workbooks with different schemas get
// Smart Data fields in Sheets; a document links both, inserts their fields
// (panel and action registry) plus a local field edited in the page, saves;
// the bound cell then moves in Sheets (row insert + cut/paste) and changes;
// the reopened document flags the link stale, updates from source and saves.
// Usage: node docs-smart-data.mjs <outDir> <DVH Office.exe>
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
const UI = { caller: 'ui', docId: 'driver', permissions: null }

// ---------- fixtures ----------
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')

async function writeXlsx(path, rows) {
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
  const cell = (ref, v) =>
    typeof v === 'number'
      ? `<c r="${ref}"><v>${v}</v></c>`
      : `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`
  const body = Object.entries(rows)
    .map(
      ([r, cells]) =>
        `<row r="${r}">${Object.entries(cells)
          .map(([col, v]) => cell(`${col}${r}`, v))
          .join('')}</row>`,
    )
    .join('')
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`,
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
    .map((t) => `<w:p><w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`)
    .join('')
  zip.file(
    'word/document.xml',
    `${XML}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${ps}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

const acceptance = join(outDir, 'p1-nghiem-thu.xlsx')
const materials = join(outDir, 'p1-vat-tu.xlsx')
const docx = join(outDir, 'p1-bao-cao.docx')
await writeXlsx(acceptance, {
  5: { A: 'Tên dự án', B: 'Dự án A' },
  6: { A: 'Giá trị HĐ', B: 1250000 },
})
await writeXlsx(materials, { 3: { A: 'Vật tư', B: 'Xi măng PCB40', C: 125.5 } })
await writeDocx(docx, [
  'Báo cáo nghiệm thu',
  'Tên dự án: ',
  'Giá trị hợp đồng: ',
  'Vật tư: ',
  'Khối lượng: ',
  'Người lập: ',
])

// ---------- file inspection ----------
async function parts(path) {
  const z = await JSZip.loadAsync(readFileSync(path))
  const items = {}
  for (const p of Object.keys(z.files).filter((p) => /^customXml\/item\d+\.xml$/.test(p))) {
    items[p] = await z.file(p).async('string')
  }
  const model = Object.values(items).find((x) => x.includes('urn:dvh-office:model:1')) ?? null
  const history = Object.values(items).find((x) => x.includes('urn:dvh-office:history:1')) ?? null
  return { z, model, history }
}
const modelFields = (model) =>
  model
    ? Object.fromEntries(
        [...model.matchAll(/<dvh:f\b([^>]*)>([^<]*)</g)].map((m) => [
          /name="([^"]*)"/.exec(m[1])?.[1],
          m[2],
        ]),
      )
    : null
const historyActions = (history) =>
  history ? [...history.matchAll(/"action":"([^"]+)"/g)].map((m) => m[1]) : []

async function workbookState(path) {
  const { z, model, history } = await parts(path)
  const wb = await z.file('xl/workbook.xml').async('string')
  return {
    names: [...wb.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)].map(
      (m) => `${/name="([^"]*)"/.exec(m[1])[1]}=${m[2]}`,
    ),
    fields: modelFields(model),
    history: historyActions(history),
  }
}

async function documentState(path) {
  const { z, model, history } = await parts(path)
  const doc = await z.file('word/document.xml').async('string')
  const sdts = [...doc.matchAll(/<w:sdt>([\s\S]*?)<\/w:sdt>/g)].map((m) => ({
    tag: /w:tag w:val="([^"]*)"/.exec(m[1])?.[1],
    bound: m[1].includes('w:dataBinding'),
    text: (/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/.exec(m[1])?.[1] ?? '').replace(
      /<[^>]+>/g,
      '',
    ),
  }))
  return {
    sdts,
    fields: modelFields(model),
    links: model
      ? JSON.parse(/<!\[CDATA\[([\s\S]*?)\]\]>/.exec(model)?.[1] ?? '{"links":[]}').links.map(
          (l) => `${l.source.relPath} rev ${l.lastSync?.revision}`,
        )
      : [],
    history: historyActions(history),
  }
}

// ---------- app driving ----------
async function withApp(path, kind, fn) {
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
        .find((p) => p.url().includes(`://${kind}/`))
      if (!page) await sleep(200)
    }
    const errors = []
    page.on('pageerror', (e) => errors.push(e.message))
    if (kind === 'sheets') {
      await page.waitForFunction(() => window.__genofficeDebug?.dvhActions, null, {
        timeout: 60000,
      })
    } else {
      await page.locator('.ProseMirror').first().waitFor({ timeout: 30000 })
      await page.waitForFunction(() => window.__dvhActions, null, { timeout: 30000 })
    }
    await sleep(2500)
    const result = await fn({ app, page })
    return { ...result, pageErrors: errors }
  } finally {
    await app.close().catch(() => {})
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
    } catch {}
  }
}

const sheetsAction = (page, name, input) =>
  page.evaluate(
    async ({ name, input, ui }) => {
      const r = await window.__genofficeDebug.dvhActions.run(name, input, {
        ...ui,
        permissions: new Set(['read', 'write']),
      })
      return r.output ?? null
    },
    { name, input, ui: UI },
  )
const docsAction = (page, name, input) =>
  page.evaluate(
    async ({ name, input, ui }) => {
      const r = await window.__dvhActions.run(name, input, {
        ...ui,
        permissions: new Set(['read', 'write']),
      })
      return JSON.parse(JSON.stringify(r.output ?? null))
    },
    { name, input, ui: UI },
  )

async function saveSheets(page) {
  await page.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(4000)
  return page.evaluate(() => document.querySelector('.status-msg')?.textContent?.trim() ?? '')
}

/** Places the caret at the end of the text node containing `needle`. */
async function caretAtEnd(page, needle) {
  const ok = await page.evaluate((needle) => {
    const root = document.querySelector('.ProseMirror')
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (!node.textContent.includes(needle)) continue
      root.focus()
      const range = document.createRange()
      range.setStart(node, node.textContent.length)
      range.collapse(true)
      const sel = window.getSelection()
      sel.removeAllRanges()
      sel.addRange(range)
      return true
    }
    return false
  }, needle)
  if (!ok) throw new Error(`text not found: ${needle}`)
  await sleep(300)
}

const stubPicker = (app, path) =>
  app.evaluate(({ dialog }, p) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [p] })
  }, path)

async function openPanel(page) {
  await page.getByText('Insert', { exact: true }).first().click()
  await page.locator('button', { hasText: 'Smart Data' }).first().click()
  const panel = page.getByRole('dialog', { name: 'Smart Data fields' })
  await panel.waitFor()
  await sleep(1500)
  return panel
}

const fieldTexts = (page) =>
  page.evaluate(() =>
    [...document.querySelectorAll('[data-dvh-field]')].map((el) => el.textContent),
  )

// ---------- 1. Sheets: fields in two workbooks with different schemas ----------
report.sheets1 = await withApp(acceptance, 'sheets', async ({ page }) => {
  await sheetsAction(page, 'Spreadsheet.BindField', { name: 'Project.Name', cell: 'B5' })
  await sheetsAction(page, 'Spreadsheet.BindField', { name: 'Contract.Value', cell: 'B6' })
  const listed = await sheetsAction(page, 'Data.ListFields', {})
  return { listed, status: await saveSheets(page) }
})
report.sheets1.file = await workbookState(acceptance)
report.sheets2 = await withApp(materials, 'sheets', async ({ page }) => {
  await sheetsAction(page, 'Spreadsheet.BindField', { name: 'Material.Name', cell: 'B3' })
  await sheetsAction(page, 'Spreadsheet.BindField', { name: 'Material.Qty', cell: 'C3' })
  return { status: await saveSheets(page) }
})
report.sheets2.file = await workbookState(materials)

// ---------- 2. Docs: link both, insert fields, edit a local field, save ----------
report.docs1 = await withApp(docx, 'docs', async ({ app, page }) => {
  const panel = await openPanel(page)
  await stubPicker(app, acceptance)
  await panel.getByRole('button', { name: 'Link workbook…' }).click()
  await sleep(1500)
  await stubPicker(app, materials)
  await panel.getByRole('button', { name: 'Link workbook…' }).click()
  await sleep(1500)
  const afterLink = await panel.innerText()

  // panel insert at the caret
  await caretAtEnd(page, 'Tên dự án:')
  await panel
    .locator('tr', { hasText: 'Project.Name' })
    .getByRole('button', { name: 'Insert' })
    .click()
  await sleep(500)
  await caretAtEnd(page, 'Giá trị hợp đồng:')
  await panel
    .locator('tr', { hasText: 'Contract.Value' })
    .getByRole('button', { name: 'Insert' })
    .click()
  await sleep(500)
  // action registry (the path the agent and scripts use)
  await docsAction(page, 'Document.InsertField', { field: 'Material.Name', blockIndex: 3 })
  await docsAction(page, 'Document.InsertField', { field: 'Material.Qty', blockIndex: 4 })
  // a document-local field, then typed into (two-way)
  await caretAtEnd(page, 'Người lập:')
  await panel.getByRole('textbox', { name: 'Field name, e.g. Project.Name' }).fill('Report.Author')
  await panel.getByRole('textbox', { name: 'Value' }).fill('Nguyễn Văn A')
  await panel.getByRole('button', { name: 'Add and insert' }).click()
  await sleep(500)
  const typed = await page.evaluate(() => {
    const el = [...document.querySelectorAll('[data-dvh-field]')].find(
      (e) => e.textContent === 'Nguyễn Văn A',
    )
    const node = el?.firstChild
    if (!node) return false
    const range = document.createRange()
    range.setStart(node, 'Nguyễn Văn'.length)
    range.collapse(true)
    document.querySelector('.ProseMirror').focus()
    const sel = window.getSelection()
    sel.removeAllRanges()
    sel.addRange(range)
    return true
  })
  await sleep(300)
  if (typed) await page.keyboard.type(' B.')
  await sleep(500)
  const fields = await fieldTexts(page)
  const listed = await docsAction(page, 'Data.ListFields', {})
  await page.screenshot({ path: join(outDir, 'docs-panel.png') })
  await page.keyboard.press('Control+s')
  await sleep(4000)
  return { afterLink, fields, listed }
})
report.docs1.file = await documentState(docx)

// ---------- 3. Sheets: move the bound cell B5 → D10, then change it ----------
report.sheets3 = await withApp(acceptance, 'sheets', async ({ page }) => {
  const ws = (code) =>
    page.evaluate(
      `(() => { const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet(); ${code} })()`,
    )
  const goto = (range) =>
    page.evaluate(
      (r) => window.__genofficeControl({ cmd: 'goto', target: { kind: 'range', range: r } }),
      range,
    )
  await ws('ws.insertRowsBefore(1, 5)')
  await sleep(800)
  const box = await page.evaluate(() => {
    const big = [...document.querySelectorAll('canvas')]
      .map((c) => c.getBoundingClientRect())
      .sort((a, b) => b.width * b.height - a.width * a.height)[0]
    return big ? { x: big.left + 120, y: big.top + 300 } : null
  })
  if (box) await page.mouse.click(box.x, box.y)
  await goto('B10')
  await page.keyboard.press('Control+x')
  await sleep(400)
  await goto('D10')
  await page.keyboard.press('Control+v')
  await sleep(1000)
  const field = await sheetsAction(page, 'Data.SetField', {
    field: 'Project.Name',
    value: 'Dự án B',
  })
  await sleep(800)
  const listed = await sheetsAction(page, 'Data.ListFields', {})
  return { field, listed, status: await saveSheets(page) }
})
report.sheets3.file = await workbookState(acceptance)

// ---------- 4. Docs: stale warning, update from source, save ----------
report.docs2 = await withApp(docx, 'docs', async ({ page }) => {
  const panel = await openPanel(page)
  await sleep(1500)
  const stale = await panel
    .locator('[data-status]')
    .evaluateAll((els) =>
      els.map(
        (e) => `${e.closest('[data-link-id]')?.textContent?.slice(0, 24)}: ${e.dataset.status}`,
      ),
    )
  const before = await fieldTexts(page)
  const link = panel.locator('[data-link-id]', { has: page.locator('[data-status="stale"]') })
  await link.getByRole('button', { name: 'Update from source' }).click()
  await sleep(1200)
  const after = await fieldTexts(page)
  await page.locator('summary', { hasText: 'History' }).click()
  await sleep(300)
  const history = await panel
    .locator('.dvh-docs-panel-history li')
    .evaluateAll((els) => els.slice(0, 6).map((e) => e.textContent.replace(/\s+/g, ' ').trim()))
  await page.screenshot({ path: join(outDir, 'docs-update.png') })
  await page.keyboard.press('Control+s')
  await sleep(4000)
  return { stale, before, after, history }
})
report.docs2.file = await documentState(docx)

writeFileSync(join(outDir, 'report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

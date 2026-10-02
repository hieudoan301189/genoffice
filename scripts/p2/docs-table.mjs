// P2 acceptance (Docs ← Sheets): a workbook range becomes a collection; a
// document links the workbook and renders the collection as DVH tables (one
// by the action registry in source style, one from the panel switched to the
// destination style); rows are added to the source in Sheets; the reopened
// document updates from source and both tables follow. Word reads the result.
// Usage: node docs-table.mjs <outDir> <DVH Office.exe>
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
const step = (label) => console.error(`${new Date().toISOString()} ${label}`)
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')

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
    .map((t) => `<w:p><w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`)
    .join('')
  zip.file(
    'word/document.xml',
    `${XML}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${ps}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  )
  writeFileSync(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

const xlsx = join(outDir, 'p2-khoi-luong.xlsx')
const docx = join(outDir, 'p2-bao-cao.docx')
await writeXlsx(xlsx)
await writeDocx(docx, ['Bảng khối lượng nghiệm thu', 'Bảng theo kiểu đích:', 'Hết.'])

async function documentState() {
  const z = await JSZip.loadAsync(readFileSync(docx))
  const doc = await z.file('word/document.xml').async('string')
  let model = null
  for (const p of Object.keys(z.files).filter((p) => /^customXml\/item\d+\.xml$/.test(p))) {
    const x = await z.file(p).async('string')
    if (x.includes('urn:dvh-office:model:1')) model = x
  }
  const objects = model ? JSON.parse(/<!\[CDATA\[([\s\S]*?)\]\]>/.exec(model)?.[1] ?? '{}') : {}
  return {
    tables: [
      ...doc.matchAll(
        /<w:sdt>(<w:sdtPr>[\s\S]*?<\/w:sdtPr>)<w:sdtContent>(<w:tbl>[\s\S]*?<\/w:tbl>)<\/w:sdtContent><\/w:sdt>/g,
      ),
    ].map((m) => ({
      tag: /w:tag w:val="([^"]*)"/.exec(m[1])?.[1],
      rows: (m[2].match(/<w:tr[ >]/g) ?? []).length,
      tblHeader: m[2].includes('<w:tblHeader/>'),
      firstFill: /<w:shd [^>]*w:fill="([^"]*)"/.exec(m[2])?.[1] ?? null,
      lastRowText: [...m[2].matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)]
        .pop()?.[0]
        .match(/<w:t[^>]*>[^<]*<\/w:t>/g)
        ?.map((t) => t.replace(/<[^>]+>/g, '')),
    })),
    collections: objects.collections?.map((c) => `${c.name}: ${c.rows.length} rows`) ?? [],
    tableModels: objects.tables?.map((t) => `${t.name} (${t.style.mode})`) ?? [],
  }
}

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
    page.on('dialog', (d) => d.accept())
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
    // a close blocked by a dialog (unsaved prompt, main-process error) must not hang the run
    const closed = await Promise.race([
      app.close().then(
        () => true,
        () => true,
      ),
      sleep(20000).then(() => false),
    ])
    if (!closed) (report.closeTimeouts ??= []).push(new Date().toISOString())
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
    } catch {}
  }
}

const run = (page, registry, name, input) =>
  page.evaluate(
    async ({ registry, name, input }) => {
      const reg = registry === 'sheets' ? window.__genofficeDebug.dvhActions : window.__dvhActions
      try {
        const r = await reg.run(name, input, {
          caller: 'ui',
          docId: 'driver',
          permissions: new Set(['read', 'write']),
        })
        return JSON.parse(JSON.stringify(r.output ?? null))
      } catch (e) {
        return { error: e.message }
      }
    },
    { registry, name, input },
  )
const ws = (page, code) =>
  page.evaluate(
    `(() => { const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet(); ${code} })()`,
  )
const docTables = (page) =>
  page.evaluate(() =>
    [...document.querySelectorAll('.ProseMirror table[data-dvh-table]')].map((t) =>
      [...t.rows].map((r) => [...r.cells].map((c) => c.textContent.trim())),
    ),
  )
async function openPanel(page) {
  await page.getByText('Insert', { exact: true }).first().click()
  await page.locator('button', { hasText: 'Smart Data' }).first().click()
  const panel = page.getByRole('dialog', { name: 'Smart Data fields' })
  await panel.waitFor()
  await sleep(1500)
  return panel
}

// 1. Sheets: a styled range becomes the WorkItems collection
report.sheets1 = await withApp(xlsx, 'sheets', async ({ page }) => {
  await ws(
    page,
    `const head = { bl: 1, bg: { rgb: '#FFF2CC' }, bd: { b: { s: 1, cl: { rgb: '#808080' } } } }
     const qty = { cl: { rgb: '#C00000' }, n: { pattern: '#,##0.00' } }
     ws.getRange('A1:C4').setValues([
       [{ v: 'Mã', s: head }, { v: 'Công việc', s: head }, { v: 'Khối lượng', s: head }],
       [{ v: 'AB.1' }, { v: 'Đào móng' }, { v: 120.5, s: qty }],
       [{ v: 'AB.2' }, { v: 'Bê tông lót' }, { v: 18, s: qty }],
       [{ v: 'AB.3' }, { v: 'Cốt thép' }, { v: 2450.25, s: qty }],
     ])`,
  )
  await sleep(500)
  const collection = await run(page, 'sheets', 'Data.CreateCollection', {
    name: 'WorkItems',
    range: 'A1:C4',
  })
  await page.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(4000)
  return { collection }
})

// 2. Docs: link, insert two tables, switch the second one to the destination style, save
report.docs1 = await withApp(docx, 'docs', async ({ app, page }) => {
  step('docs1: panel')
  const panel = await openPanel(page)
  step('docs1: stub picker')
  await app.evaluate(({ dialog }, p) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [p] })
  }, xlsx)
  step('docs1: link')
  await panel.getByRole('button', { name: 'Link workbook…' }).click()
  await sleep(1500)
  step('docs1: insert via action')
  const viaAction = await run(page, 'docs', 'Document.InsertTable', {
    collection: 'WorkItems',
    afterBlockIndex: 0,
  })
  await sleep(500)
  // panel insert after the "Bảng theo kiểu đích:" paragraph (caret there)
  await page.evaluate(() => {
    const root = document.querySelector('.ProseMirror')
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (!node.textContent.includes('kiểu đích')) continue
      root.focus()
      const range = document.createRange()
      range.setStart(node, node.textContent.length)
      range.collapse(true)
      const sel = window.getSelection()
      sel.removeAllRanges()
      sel.addRange(range)
      return
    }
  })
  await sleep(300)
  step('docs1: insert via panel')
  await page.screenshot({ path: join(outDir, 'docs-before-panel-insert.png') })
  await panel
    .locator('tr', { hasText: 'WorkItems' })
    .getByRole('button', { name: 'Insert table' })
    .click()
  await sleep(800)
  step('docs1: mode select')
  const tables = page.getByRole('table', { name: 'Tables' })
  await tables.locator('select').nth(1).selectOption('destination')
  await sleep(800)
  step('docs1: read tables')
  const rendered = await docTables(page)
  await page.screenshot({ path: join(outDir, 'docs-tables.png') })
  // Ctrl+S goes to the editor, not to a panel control that kept focus
  await page
    .locator('.ProseMirror')
    .first()
    .click({ position: { x: 5, y: 5 } })
  await page.keyboard.press('Control+s')
  await sleep(4000)
  return { viaAction, rendered }
})
report.docs1.file = await documentState()

// 3. Sheets: two rows appended under the collection range
report.sheets2 = await withApp(xlsx, 'sheets', async ({ page }) => {
  await ws(
    page,
    `ws.getRange('A5:C6').setValues([['AB.4', 'Ván khuôn', 75], ['AB.5', 'Xây gạch', 32.4]])`,
  )
  await sleep(500)
  await page.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(4000)
  return {}
})

// 4. Docs: the link is stale; update from source refreshes both tables
report.docs2 = await withApp(docx, 'docs', async ({ page }) => {
  const panel = await openPanel(page)
  await sleep(1000)
  const stale = await panel
    .locator('[data-status]')
    .evaluateAll((els) => els.map((e) => e.dataset.status))
  const before = await docTables(page)
  await panel.getByRole('button', { name: 'Update from source' }).click()
  await sleep(1500)
  const after = await docTables(page)
  await page.screenshot({ path: join(outDir, 'docs-tables-updated.png') })
  // Ctrl+S goes to the editor, not to a panel control that kept focus
  await page
    .locator('.ProseMirror')
    .first()
    .click({ position: { x: 5, y: 5 } })
  await page.keyboard.press('Control+s')
  await sleep(4000)
  return { stale, before: before.map((t) => t.length), after }
})
report.docs2.file = await documentState()

writeFileSync(join(outDir, 'report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

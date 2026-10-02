// P2 acceptance (Sheets): a styled range becomes a collection, a DVH.Table
// renders it elsewhere on the sheet, follows added/removed source rows
// (content below moves), asks before overwriting a hand edit, switches style
// mode, renders 1,000 rows, and survives save + reopen.
// Usage: node sheets-table.mjs <outDir> <DVH Office.exe>
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
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

async function writeBlankXlsx(path) {
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

const path = join(outDir, 'p2-sheets.xlsx')
await writeBlankXlsx(path)

async function fileState() {
  const z = await JSZip.loadAsync(readFileSync(path))
  const wb = await z.file('xl/workbook.xml').async('string')
  let model = null
  for (const p of Object.keys(z.files).filter((p) => /^customXml\/item\d+\.xml$/.test(p))) {
    const x = await z.file(p).async('string')
    if (x.includes('urn:dvh-office:model:1')) model = x
  }
  const objects = model ? JSON.parse(/<!\[CDATA\[([\s\S]*?)\]\]>/.exec(model)?.[1] ?? '{}') : null
  return {
    names: [...wb.matchAll(/<definedName\b([^>]*)>([^<]*)<\/definedName>/g)].map(
      (m) => `${/name="([^"]*)"/.exec(m[1])[1]}=${m[2]}`,
    ),
    collections: objects?.collections?.map((c) => `${c.name}: ${c.rows.length} rows`) ?? [],
    tables: objects?.tables?.map((tb) => `${tb.name} (${tb.style.mode})`) ?? [],
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
    const errors = []
    page.on('pageerror', (e) => errors.push(e.message))
    page.on('dialog', (d) => d.accept())
    await page.waitForFunction(() => window.__genofficeDebug?.dvhActions, null, { timeout: 60000 })
    await sleep(2500)
    const result = await fn(page)
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

const action = (page, name, input) =>
  page.evaluate(
    async ({ name, input }) => {
      try {
        const r = await window.__genofficeDebug.dvhActions.run(name, input, {
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
const ws = (page, code) =>
  page.evaluate(
    `(() => { const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet(); ${code} })()`,
  )
const values = (page, a1) => ws(page, `return ws.getRange('${a1}').getValues()`)
const styleAt = (page, a1) =>
  ws(
    page,
    `const s = ws.getRange('${a1}').getCellStyleData() ?? {}; return { bl: s.bl ?? 0, bg: s.bg?.rgb ?? null, cl: s.cl?.rgb ?? null, n: s.n?.pattern ?? null, bd: !!s.bd?.b }`,
  )
async function save(page) {
  await page.getByRole('button', { name: /^Save \(/ }).click()
  await sleep(4000)
}

report.session1 = await withApp(async (page) => {
  // a styled source range: yellow bold titles, red quantities with a number format
  await ws(
    page,
    `const head = { bl: 1, bg: { rgb: '#FFF2CC' }, bd: { b: { s: 1, cl: { rgb: '#808080' } } } }
     const qty = { cl: { rgb: '#C00000' }, n: { pattern: '#,##0.00' } }
     ws.getRange('A1:C4').setValues([
       [{ v: 'Mã', s: head }, { v: 'Công việc', s: head }, { v: 'Khối lượng', s: head }],
       [{ v: 'AB.1' }, { v: 'Đào móng' }, { v: 120.5, s: qty }],
       [{ v: 'AB.2' }, { v: 'Bê tông lót' }, { v: 18, s: qty }],
       [{ v: 'AB.3' }, { v: 'Cốt thép' }, { v: 2450.25, s: qty }],
     ])
     ws.getRange('E12').setValue('Ghi chú dưới bảng')`,
  )
  await sleep(500)
  const out = {}
  out.collection = await action(page, 'Data.CreateCollection', {
    name: 'WorkItems',
    range: 'A1:C4',
  })
  out.table = await action(page, 'Table.Create', { name: 'Bảng KL', collection: 'WorkItems' })
  out.render = await action(page, 'Table.Render', { table: 'Bảng KL', cell: 'E2' })
  await sleep(800)
  out.rendered = await values(page, 'E2:G5')
  out.headerStyle = await styleAt(page, 'E2')
  out.qtyStyle = await styleAt(page, 'G3')

  // a new source row typed under the range joins the collection; the table grows
  await ws(page, `ws.getRange('A5:C5').setValues([['AB.4', 'Ván khuôn', 75]])`)
  await sleep(500)
  out.grow = await action(page, 'Table.Refresh', { table: 'Bảng KL' })
  await sleep(800)
  out.afterGrow = { table: await values(page, 'E2:G6'), noteE13: await values(page, 'E13') }

  // two rows removed from the source; the table shrinks and the note moves back
  await ws(page, `ws.getRange('A4:C5').setValues([[null, null, null], [null, null, null]])`)
  await sleep(500)
  out.shrink = await action(page, 'Table.Refresh', { table: 'Bảng KL' })
  await sleep(800)
  out.afterShrink = { table: await values(page, 'E2:G5'), noteE11: await values(page, 'E11') }

  // hand edit inside the table: refresh refuses, force overwrites
  await ws(page, `ws.getRange('F3').setValue('sửa tay')`)
  await sleep(500)
  out.refreshEdited = await action(page, 'Table.Refresh', { table: 'Bảng KL' })
  out.refreshForced = await action(page, 'Table.Refresh', { table: 'Bảng KL', force: true })
  await sleep(500)
  out.afterForce = await values(page, 'F3')

  // destination style
  await action(page, 'Table.SetStyle', {
    table: 'Bảng KL',
    mode: 'destination',
    bandFill: '#F2F2F2',
  })
  out.destination = await action(page, 'Table.Refresh', { table: 'Bảng KL' })
  await sleep(500)
  out.destinationHeader = await styleAt(page, 'E2')
  out.destinationBand = await styleAt(page, 'E4')

  // 1,000 rows x 10 columns: collection + render timing
  out.big = await page.evaluate(async () => {
    const ws = window.__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
    if (ws.getMaxRows() < 1100) ws.insertRowsAfter(ws.getMaxRows() - 1, 1100 - ws.getMaxRows())
    const head = Array.from({ length: 10 }, (_, c) => ({ v: `Cột ${c + 1}`, s: { bl: 1 } }))
    const body = Array.from({ length: 1000 }, (_, r) =>
      Array.from({ length: 10 }, (_, c) => ({ v: c === 0 ? `R${r + 1}` : r * 10 + c })),
    )
    ws.getRange(20, 0, 1001, 10).setValues([head, ...body])
    await new Promise((r) => setTimeout(r, 500))
    const run = (name, input) =>
      window.__genofficeDebug.dvhActions.run(name, input, {
        caller: 'ui',
        docId: 'driver',
        permissions: new Set(['read', 'write']),
      })
    await run('Data.CreateCollection', { name: 'Big', range: 'A21:J1021' })
    await run('Table.Create', { name: 'Big table', collection: 'Big', mode: 'destination' })
    const t0 = performance.now()
    const r = await run('Table.Render', { table: 'Big table', cell: 'L21' })
    const ms = performance.now() - t0
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))
    return { ms: Math.round(ms), painted: Math.round(performance.now() - t0), output: r.output }
  })
  out.bigLast = await values(page, 'L1021:U1021')
  await save(page)
  return out
})
report.saved = await fileState()

report.session2 = await withApp(async (page) => {
  const out = {}
  out.listed = await page.evaluate(() =>
    window.__genofficeDebug.univerAPI
      .getActiveWorkbook()
      .getDefinedNames()
      .map((n) => `${n.getName()}=${n.getFormulaOrRefString()}`)
      .filter((n) => n.startsWith('_dvh.')),
  )
  // no hand edits since the save: refresh must not ask
  out.refreshAfterReopen = await action(page, 'Table.Refresh', { table: 'Bảng KL' })
  await page.getByRole('button', { name: 'Smart Data', exact: true }).click()
  const panel = page.getByRole('dialog', { name: 'Smart Data fields' })
  await panel.waitFor()
  await sleep(1200)
  out.panel = await panel.innerText()
  await page.screenshot({ path: join(outDir, 'panel.png') })
  return out
})

writeFileSync(join(outDir, 'report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

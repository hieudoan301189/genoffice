// Spike S2: a Word-saved docx with bound DVH Smart Fields is edited in the
// packaged DVH Office Docs editor — text typed before a field, after a field,
// and inside a field — then saved; the controls and their data binding must
// survive. Usage: node app-check.mjs <outDir> <word-saved docx> <DVH Office.exe>
import { _electron as electron } from 'playwright-core'
import JSZip from 'jszip'
import { copyFileSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { execFileSync } from 'node:child_process'

const outDir = resolve(process.argv[2])
const source = resolve(process.argv[3])
const exe = resolve(process.argv[4])
mkdirSync(outDir, { recursive: true })
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const docxPath = join(outDir, 's2-edited.docx')
copyFileSync(source, docxPath)
const report = {}

/** Places the caret `offset` characters into the first text node containing `needle`. */
async function caretAt(page, needle, offset) {
  const ok = await page.evaluate(
    ({ needle, offset }) => {
      const root = document.querySelector('.ProseMirror')
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const at = node.textContent.indexOf(needle)
        if (at < 0) continue
        root.focus()
        const range = document.createRange()
        range.setStart(node, at + offset)
        range.collapse(true)
        const sel = window.getSelection()
        sel.removeAllRanges()
        sel.addRange(range)
        return true
      }
      return false
    },
    { needle, offset },
  )
  if (!ok) throw new Error(`text not found: ${needle}`)
  await sleep(300)
}

const profile = join(outDir, `profile-${Date.now()}`)
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
  await home.evaluate((p) => window.aiOffice.openPath(p), docxPath)
  let page
  for (let i = 0; i < 150 && !page; i++) {
    page = app
      .context()
      .pages()
      .find((p) => p.url().includes('://docs/'))
    if (!page) await sleep(200)
  }
  const errors = []
  page.on('pageerror', (e) => errors.push(e.message))
  await page.locator('.ProseMirror').first().waitFor({ timeout: 30000 })
  await sleep(3000)
  report.fieldSpans = await page.evaluate(() =>
    [...document.querySelectorAll('[data-dvh-field]')].map((el) => el.textContent),
  )

  // 1. before the project field, in the same paragraph
  await caretAt(page, 'Tên dự án', 0)
  await page.keyboard.type('Ghi chú – ')
  // 2. right after the date field (non-inclusive mark: stays outside)
  await caretAt(page, '01/10/2026', 10)
  await page.keyboard.type(' (dự kiến)')
  // 3. inside the project field
  await caretAt(page, 'Gói thầu số 7', 'Gói thầu số'.length)
  await page.keyboard.type(' A')
  await sleep(500)
  report.fieldSpansAfterTyping = await page.evaluate(() =>
    [...document.querySelectorAll('[data-dvh-field]')].map((el) => el.textContent),
  )
  await page.screenshot({ path: join(outDir, 's2-editor.png') })
  await page.keyboard.press('Control+s')
  await sleep(3500)
  report.pageErrors = errors
} finally {
  await app.close().catch(() => {})
  try {
    execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' })
  } catch {}
}

const before = await JSZip.loadAsync(readFileSync(source))
const after = await JSZip.loadAsync(readFileSync(docxPath))
const doc = await after.file('word/document.xml').async('string')
const custom = async (zip) =>
  Object.fromEntries(
    await Promise.all(
      Object.keys(zip.files)
        .filter((p) => p.startsWith('customXml/') && !zip.files[p].dir)
        .map(async (p) => [p, await zip.file(p).async('string')]),
    ),
  )
const [cb, ca] = [await custom(before), await custom(after)]
const paragraphOf = (needle) => {
  const at = doc.indexOf(needle)
  if (at < 0) return null
  const start =
    doc.lastIndexOf('<w:p>', at) >= 0
      ? Math.max(doc.lastIndexOf('<w:p>', at), doc.lastIndexOf('<w:p ', at))
      : 0
  return doc.slice(start, doc.indexOf('</w:p>', at) + 6)
}
const projectPara = paragraphOf('Ghi chú')
const datePara = paragraphOf('(dự kiến)')
report.saved = {
  customXmlIdentical: JSON.stringify(cb) === JSON.stringify(ca),
  sdtCount: (doc.match(/<w:sdt>/g) ?? []).length,
  projectParagraph: {
    typedBeforeOutsideControl: /Ghi chú – [\s\S]*?<w:sdt>/.test(projectPara ?? ''),
    tagAndBinding: /w:tag w:val="dvh:f:f_projectname0001"[\s\S]*w:dataBinding/.test(
      projectPara ?? '',
    ),
    controlText: /<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/
      .exec(projectPara ?? '')?.[1]
      ?.replace(/<[^>]+>/g, ''),
  },
  dateParagraph: {
    tagAndBinding: /w:tag w:val="dvh:f:f_acceptdate00001"[\s\S]*w:dataBinding/.test(datePara ?? ''),
    controlText: /<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/
      .exec(datePara ?? '')?.[1]
      ?.replace(/<[^>]+>/g, ''),
    typedAfterOutsideControl: /<\/w:sdt>[\s\S]*\(dự kiến\)/.test(datePara ?? ''),
  },
  tableSdt: doc.includes('w:tag w:val="dvh:t:t_workitems000001"'),
}
writeFileSync(join(outDir, 'app-report.json'), JSON.stringify(report, null, 2))
console.log(JSON.stringify(report, null, 2))

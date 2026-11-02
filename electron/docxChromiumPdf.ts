import { BrowserWindow } from 'electron'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { paginateFlowSections } from '../core/paginateFlow'
import { tightenLineHeight } from '../core/docxScale'

/**
 * The app's docx → PDF converter. Renders the .docx with docx-preview
 * (the same library the on-screen preview uses, with the repo's patches) inside a hidden Chromium
 * window, then prints that page with Electron's own printToPDF. No external program is needed.
 *
 * No external program (LibreOffice or otherwise) is needed.
 */
export function convertDocxToPdfWithChromium(docxBuffer: Buffer): Promise<Buffer> {
  // One conversion at a time on one shared window: a fresh BrowserWindow per call
  // intermittently fails its second load with ERR_FAILED (same Electron quirk
  // documented in pdfRaster.ts), while sequential navigations of one window don't.
  const run = queue.then(() => convert(docxBuffer))
  queue = run.catch(() => undefined)
  return run
}

let queue: Promise<unknown> = Promise.resolve()
let win: BrowserWindow | null = null

function getWindow(): BrowserWindow {
  if (win && !win.isDestroyed()) return win
  win = new BrowserWindow({
    show: false,
    width: 900,
    height: 1200,
    webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false }
  })
  return win
}

// A long document whose last page holds only a line or two (the Civil Tender Document's signature line spilling onto page 42
// when a long work name adds lines upstream) is tightened by the smallest step that pulls it back; if no step does, the
// original layout is kept.
const ORPHAN_MIN_SHEETS = 8
const ORPHAN_MAX_CHARS = 160
// Line height only, so the font size (and so how the lines wrap) never changes.
const SHRINK_STEPS = [0.995, 0.99, 0.985, 0.98, 0.97, 0.96, 0.95]

async function convert(docxBuffer: Buffer): Promise<Buffer> {
  const first = await convertOnce(docxBuffer)
  if (!first.orphan) return first.pdf
  for (const factor of SHRINK_STEPS) {
    const tightened = Buffer.from(tightenLineHeight(docxBuffer.toString('base64'), factor), 'base64')
    const next = await convertOnce(tightened)
    if (next.sheets < first.sheets) return next.pdf
  }
  return first.pdf
}

async function convertOnce(docxBuffer: Buffer): Promise<{ pdf: Buffer; sheets: number; orphan: boolean }> {
  // Packaged builds keep these two unpacked (electron-builder.yml asarUnpack) so the page can load them by file:// URL.
  const unpacked = (p: string) => p.replace('app.asar' + path.sep, 'app.asar.unpacked' + path.sep)
  const jszipPath = unpacked(require.resolve('jszip/dist/jszip.min.js'))
  const docxPreviewPath = unpacked(require.resolve('docx-preview'))

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'docugen-docx2pdf-'))
  // Deliberately no <!DOCTYPE>: quirks mode lays Word line boxes out the way Word does (the 41-page Tender Document
  // comes out at 41 pages; standards mode runs it to ~52 with taller lines).
  const page = path.join(dir, 'page.html')
  fs.writeFileSync(
    page,
    `<html><head><meta charset="utf-8">
<style>
  @page { size: A4; margin: 0 }
  html, body { margin: 0; padding: 0; background: #fff }
  .docx-wrapper { background: #fff !important; padding: 0 !important; margin: 0 !important }
  .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 auto !important; break-after: page }
  .docx-wrapper > section.docx:last-child { break-after: auto }
</style>
<script src="${pathToFileUrl(jszipPath)}"></script>
<script src="${pathToFileUrl(docxPreviewPath)}"></script>
</head><body><div id="root"></div></body></html>`
  )

  const w = getWindow()
  try {
    await w.loadFile(page)
    // Same render options as the on-screen preview (src/components/docPage.ts DOCX_PREVIEW_OPTIONS).
    const stats: { sheets: number; lastChars: number } = await w.webContents.executeJavaScript(`(async () => {
      const b64 = ${JSON.stringify(docxBuffer.toString('base64'))}
      const bin = atob(b64)
      const bytes = new Uint8Array(bin.length)
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
      await docx.renderAsync(bytes, document.getElementById('root'), null, {
        className: 'docx', inWrapper: true, ignoreWidth: false, ignoreHeight: false, ignoreFonts: false,
        breakPages: true, experimental: true, trimXmlDeclaration: true, useBase64URL: true,
        renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true, renderChanges: false
      })
      await document.fonts.ready
      // docx-preview sizes every tab stop (right-aligned "Date:" on the Work Order number line, etc.) in a pass it schedules
      // 500ms after renderAsync resolves. Printing before it fires leaves each tab a bare space, so right-tabbed text hugs
      // the text before it instead of sitting at the stop. Wait it out, before pagination moves paragraphs around.
      // Only for short documents (the letters with a right-tabbed "Date:"): in the 41-page Civil Tender Document the pass
      // only resizes the hanging-indent clause numbers ("8.1<tab>text"), which wraps ~16 clauses onto an extra line and
      // pushes the count to 43; with the tabs left unsized it prints at its fixed 41 pages. Its only right/center stops
      // are never the target of a tab character, so nothing in it needs them.
      const renderedHeight = [...document.querySelectorAll('.docx-wrapper > section.docx')]
        .reduce((h, s) => h + s.getBoundingClientRect().height, 0)
      if (renderedHeight < 8 * 1123) await new Promise(r => setTimeout(r, 700))
      // docx-preview stretches Word text boxes (the WordArt letterhead) to full page width and leaves curved WordArt
      // absolutely positioned off-page; same fix the on-screen preview applies (docPage.ts normalizeDocxTextboxes).
      for (const d of document.querySelectorAll('div[style*="text-indent: 0px"]')) {
        if (d.querySelector('img')) continue
        Object.assign(d.style, { width: 'auto', maxWidth: '100%', display: 'block', textAlign: 'center', transform: 'none' })
        for (const el of d.querySelectorAll('*')) {
          if (el.style.width === '100%') el.style.width = 'auto'
          if (el.closest('foreignObject') && /scale/i.test(el.style.transform)) el.style.transform = 'none'
        }
      }
      for (const svg of document.querySelectorAll('svg[style*="v-text-anchor"]')) {
        if ((svg.getAttribute('style') ?? '').includes('left:0')) {
          const parent = svg.parentElement
          if (parent && getComputedStyle(parent).position === 'static') parent.style.position = 'relative'
        } else Object.assign(svg.style, { position: 'static', display: 'block', margin: '0 auto' })
      }
      // docx-preview lays each section out as one tall sheet, which Chromium would cut at arbitrary points with no
      // margins or footer. Cut it into real page sheets (each repeating the footer) and number them, as Word does.
      const paginateFlowSections = ${paginateFlowSections.toString()}
      // A pass can leave a sheet a few px over (block margins collapse differently once moved), so go again until stable.
      for (let pass = 0; pass < 4; pass++) {
        const before = document.querySelectorAll('.docx-wrapper > section.docx').length
        paginateFlowSections(document.getElementById('root'))
        if (document.querySelectorAll('.docx-wrapper > section.docx').length === before) break
      }
      // A trailing sheet holding nothing visible (only blank paragraphs left over after the last real line) would print
      // as a blank page.
      for (;;) {
        const all = [...document.querySelectorAll('.docx-wrapper > section.docx')]
        const last = all[all.length - 1]
        if (all.length < 2 || last.textContent.trim() || last.querySelector('img, svg, table, canvas')) break
        last.remove()
      }
      const sheets = [...document.querySelectorAll('.docx-wrapper > section.docx')]
      sheets.forEach((sec, i) => {
        for (const foot of sec.querySelectorAll(':scope > footer, :scope > header')) {
          const spans = [...foot.querySelectorAll('span')]
          spans.forEach((sp, k) => {
            const next = spans[k + 1]
            if (!next || !/^\\d+$/.test(next.textContent.trim())) return
            if (/^\\s*Page\\s*$/i.test(sp.textContent)) next.textContent = String(i + 1)
            else if (/^\\s*of\\s*$/i.test(sp.textContent)) next.textContent = String(sheets.length)
          })
        }
      })
      // A document can mix page sizes (e.g. US Letter sections in an A4 file). Each section is its own printed
      // page, so shrink any wider than A4 (793.7px) to fit instead of being clipped at the edge.
      for (const sec of document.querySelectorAll('.docx-wrapper > section.docx')) {
        const w = sec.getBoundingClientRect().width
        if (w > 794) sec.style.zoom = String(793 / w)
      }
      // A sheet is min-height 1122.53px but an A4 page is 1122.52px: that hair of overflow printed a blank or
      // near-empty extra page after some sheets. Pin sheets just under A4 so they can't spill. A taller page size
      // (the SE letters are US Legal) keeps its size when its content fills it; if the content stops within A4 the
      // empty tail is cropped, otherwise the whole sheet is scaled down to fit one A4 page.
      for (const sec of document.querySelectorAll('.docx-wrapper > section.docx')) {
        const box = sec.getBoundingClientRect()
        if (box.height < 1124) {
          Object.assign(sec.style, { height: '1122px', minHeight: '1122px', overflow: 'hidden' })
          continue
        }
        const art = sec.querySelector('article')
        const used = art ? art.getBoundingClientRect().bottom - box.top + parseFloat(getComputedStyle(sec).paddingBottom) : box.height
        if (used <= 1122) Object.assign(sec.style, { height: '1122px', minHeight: '1122px', overflow: 'hidden' })
        else sec.style.zoom = String(Math.min(1, 1120 / used))
      }
      const lastSheet = sheets[sheets.length - 1]
      return { sheets: sheets.length, lastChars: (lastSheet?.querySelector('article')?.textContent ?? '').trim().length }
    })()`)
    const pdf = await w.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true })
    return { pdf, sheets: stats.sheets, orphan: stats.sheets >= ORPHAN_MIN_SHEETS && stats.lastChars < ORPHAN_MAX_CHARS }
  } finally {
    fs.rm(dir, { recursive: true, force: true }, () => {})
  }
}

function pathToFileUrl(p: string): string {
  return 'file://' + encodeURI(p.replace(/\\/g, '/').replace(/^([A-Za-z]):/, '/$1:'))
}

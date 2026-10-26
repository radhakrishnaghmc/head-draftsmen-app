import { BrowserWindow } from 'electron'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { scheduleAWorkbookToHtml } from '../core/scheduleAHtml'

/**
 * A filled Schedule-A workbook as a PDF: the sheet is drawn as HTML (see
 * core/scheduleAHtml.ts) and printed by a hidden Chromium window, the same way
 * the Word documents are (docxChromiumPdf.ts). A4 portrait, scaled to the sheet's
 * own page width, "Page n of N" footer as in the Excel file, and the print area
 * runs through the Executive / Superintending Engineer signature block.
 */
export function scheduleAWorkbookToPdf(xlsx: Buffer): Promise<Buffer> {
  // One print at a time on one shared window — see docxChromiumPdf.ts for why.
  const run = queue.then(() => print(xlsx))
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

const A4_WIDTH_IN = 8.27

/**
 * The one page the PDF, the on-screen preview and Print all draw from: the sheet
 * shrunk to the printable A4 width (CSS px are 96 per inch; never enlarged), with
 * the sheet's own page margins as @page. On screen it sits on a grey desk at A4
 * width, so what you preview is what prints.
 */
export async function scheduleAPageHtml(xlsx: Buffer): Promise<string> {
  const { html, tableWidthPx, margins } = await scheduleAWorkbookToHtml(xlsx)
  const printablePx = (A4_WIDTH_IN - margins.left - margins.right) * 96
  const zoom = Math.min(1, printablePx / tableWidthPx)
  const style = `<style>
  @page { size: A4; margin: ${margins.top}in ${margins.right}in ${margins.bottom}in ${margins.left}in }
  @media screen {
    html { background: #d8d8dc }
    body { box-sizing: border-box; width: ${A4_WIDTH_IN}in; min-height: 11.69in; margin: 12px auto; background: #fff;
      padding: ${margins.top}in ${margins.right}in ${margins.bottom}in ${margins.left}in; box-shadow: 0 0 0 1px rgba(16, 24, 40, 0.12) }
  }
</style>`
  return html.replace('</head>', `${style}</head>`).replace('<div id="sheet">', `<div id="sheet" style="zoom:${zoom}">`)
}

async function print(xlsx: Buffer): Promise<Buffer> {
  const { margins } = await scheduleAWorkbookToHtml(xlsx)
  const html = await scheduleAPageHtml(xlsx)

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'docugen-schedulea-'))
  try {
    const page = path.join(dir, 'page.html')
    fs.writeFileSync(page, html)
    const w = getWindow()
    await w.loadFile(page)
    return await w.webContents.printToPDF({
      pageSize: 'A4',
      printBackground: true,
      margins: { top: margins.top, bottom: margins.bottom, left: margins.left, right: margins.right },
      displayHeaderFooter: true,
      headerTemplate: '<span></span>',
      footerTemplate:
        '<div style="width:100%;text-align:center;font-size:8px;font-family:Arial,Helvetica,sans-serif">Page <span class="pageNumber"></span> of <span class="totalPages"></span></div>'
    })
  } finally {
    fs.rm(dir, { recursive: true, force: true }, () => {})
  }
}

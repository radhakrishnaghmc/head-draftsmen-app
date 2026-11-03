import { BrowserWindow, app } from 'electron'
import * as fs from 'fs'
import * as path from 'path'
import { excelToPrintHtml, type ExcelPrintOptions } from '../core/excelPrintHtml'

/**
 * The Excel → Printable PDF tool: the workbook is drawn as HTML (core/excelPrintHtml.ts — page size, orientation,
 * margins and scale-to-width are decided there) and printed by Chromium, the same way the other generated PDFs are.
 * The PDF and the direct print both come from that one HTML, so what is saved is what is printed.
 */
const FOOTER =
  '<div style="width:100%;text-align:center;font-size:8px;font-family:Arial,Helvetica,sans-serif">Page <span class="pageNumber"></span> of <span class="totalPages"></span></div>'

async function withPage<T>(
  xlsx: Buffer,
  opts: ExcelPrintOptions,
  visible: boolean,
  use: (w: BrowserWindow) => Promise<T>
): Promise<T> {
  const { html, sheets } = await excelToPrintHtml(xlsx, opts)
  if (sheets === 0) throw new Error('Nothing to print — the selected sheets are empty.')
  const file = path.join(app.getPath('temp'), `docugen-excelprint-${Date.now()}.html`)
  fs.writeFileSync(file, html, 'utf8')
  const w = new BrowserWindow({
    show: visible,
    width: 900,
    height: 1000,
    title: 'Print',
    webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false }
  })
  const cleanup = () => fs.rm(file, { force: true }, () => {})
  try {
    await w.loadFile(file)
    await w.webContents.executeJavaScript('document.fonts.ready.then(() => undefined)').catch(() => {})
    const result = await use(w)
    if (!visible) w.destroy()
    return result
  } catch (e) {
    if (!w.isDestroyed()) w.destroy()
    throw e
  } finally {
    if (visible) w.on('closed', cleanup)
    else cleanup()
  }
}

let queue: Promise<unknown> = Promise.resolve()

export function excelToPrintablePdf(xlsx: Buffer, opts: ExcelPrintOptions): Promise<Buffer> {
  const run = queue.then(() =>
    withPage(xlsx, opts, false, (w) =>
      w.webContents.printToPDF({
        preferCSSPageSize: true,
        printBackground: true,
        displayHeaderFooter: true,
        headerTemplate: '<span></span>',
        footerTemplate: FOOTER
      })
    )
  )
  queue = run.catch(() => undefined)
  return run
}

/** Opens the OS print dialog on the laid-out pages. The window stays open for the user to dismiss afterwards. */
export function printExcel(xlsx: Buffer, opts: ExcelPrintOptions): Promise<void> {
  const run = queue.then(() =>
    withPage(xlsx, opts, true, async (w) => {
      w.focus()
      w.webContents.print({ printBackground: true, margins: { marginType: 'default' } }, (success, reason) => {
        if (!success && reason !== 'cancelled') console.error('Excel print failed:', reason)
      })
    })
  )
  queue = run.catch(() => undefined)
  return run
}

import { BrowserWindow, screen } from 'electron'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'

/**
 * Rasterizes each page of a PDF to a PNG at an exact pixel size, using
 * Electron's own bundled Chromium/PDFium PDF viewer rather than LibreOffice's
 * `--convert-to png`. LibreOffice's rasterizer has its own bug mis-shaping
 * Telugu/complex-script glyphs during PDF→PNG export (reported as e.g.
 * "Quthbullapur" rendering as "8uthbulla/ur") — confirmed by comparing the
 * exact same PDF rendered through Poppler (correct), pdf.js (garbled by an
 * unrelated font-substitution bug — see core/docxToPdf.ts), and
 * LibreOffice's own PNG export (also garbled). PDFium renders it correctly,
 * so a hidden BrowserWindow stands in for LibreOffice at this one step.
 *
 * One BrowserWindow is created and reused across every page (navigating it
 * to a new `#page=N` fragment of the same file:// URL each time) rather than
 * a fresh window per page — a fresh window per page reliably fails with
 * `ERR_FAILED` loading the second page's URL (reproduced even with a
 * distinct temp-file copy per page, so it isn't path-based caching); reusing
 * one window's webContents for sequential navigations does not hit this.
 */
export async function rasterizePdfPages(
  pdf: Buffer,
  pageSizes: { pixelWidth: number; pixelHeight: number }[]
): Promise<Buffer[]> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'docugen-raster-'))
  const pdfPath = path.join(dir, 'doc.pdf')
  fs.writeFileSync(pdfPath, pdf)

  const scaleFactor = screen.getPrimaryDisplay().scaleFactor || 1
  const maxWidth = Math.max(...pageSizes.map((s) => s.pixelWidth))
  const maxHeight = Math.max(...pageSizes.map((s) => s.pixelHeight))

  const winWidth = Math.max(1, Math.round(maxWidth / scaleFactor))
  const winHeight = Math.max(1, Math.round(maxHeight / scaleFactor))

  // `show: false` windows never get real compositor paints on some macOS
  // setups, which starves capturePage() of pixels (see PDFium blank-capture
  // note below). Keep it `show: true` but park it almost entirely off the
  // primary display — only ~3% of its width overlaps real screen bounds —
  // so the window still counts as on-screen and gets composited, while
  // staying effectively invisible to the user.
  const display = screen.getPrimaryDisplay().bounds
  const overlapPx = Math.max(1, Math.round(winWidth * 0.03))

  const win = new BrowserWindow({
    width: winWidth,
    height: winHeight,
    x: display.x + display.width - overlapPx,
    y: display.y,
    useContentSize: true,
    show: true,
    focusable: false,
    webPreferences: { plugins: true, sandbox: true }
  })
  win.showInactive()

  try {
    const images: Buffer[] = []
    for (let i = 0; i < pageSizes.length; i++) {
      const { pixelWidth, pixelHeight } = pageSizes[i]
      await win.loadURL(`file://${pdfPath}#page=${i + 1}&toolbar=0&navpanes=0&statusbar=0&view=FitH`)
      // The PDF viewer extension paints asynchronously after loadURL resolves.
      await new Promise((resolve) => setTimeout(resolve, 400))
      const captured = await win.webContents.capturePage()
      const resized = captured.resize({ width: pixelWidth, height: pixelHeight })
      images.push(resized.toPNG())
    }
    return images
  } finally {
    win.destroy()
    fs.rmSync(dir, { recursive: true, force: true })
  }
}

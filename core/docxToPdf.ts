import { execFile } from 'child_process'
import { promisify } from 'util'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { PDFDocument } from 'pdf-lib'

const execFileAsync = promisify(execFile)

const SOFFICE_CANDIDATES: Partial<Record<NodeJS.Platform, string[]>> = {
  darwin: ['/Applications/LibreOffice.app/Contents/MacOS/soffice'],
  win32: [
    'C:/Program Files/LibreOffice/program/soffice.exe',
    'C:/Program Files (x86)/LibreOffice/program/soffice.exe'
  ],
  linux: ['/usr/bin/soffice', '/usr/bin/libreoffice', '/snap/bin/libreoffice', '/opt/libreoffice/program/soffice']
}

function findSofficeBinary(): string {
  const candidates = SOFFICE_CANDIDATES[process.platform] ?? []
  const found = candidates.find((p) => fs.existsSync(p))
  if (!found) throw new Error('Could not find the soffice binary.')
  return found
}

/**
 * Shells out to `soffice --headless --convert-to <outExt>` on one input file,
 * with no `-env:UserInstallation` override, so every call reuses LibreOffice's
 * own real, persistent user profile — the same one a normal `soffice` launch
 * (or this app's own developer terminal testing) already warms up with font
 * caches and config. (The `libreoffice-convert` package this used to go
 * through instead points `-env:UserInstallation` at a brand new, empty tmp
 * directory for every single conversion, then deletes it — so LibreOffice has
 * to rebuild its font/config cache from nothing each time, which turned out
 * not to be the source of the garbled-text bug below, but is still worth
 * avoiding on general principle.)
 */
async function sofficeConvert(
  soffice: string,
  input: Buffer,
  inExt: string,
  convertTo: string,
  outExt: string = convertTo
): Promise<Buffer> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'docugen-convert-'))
  try {
    const inputPath = path.join(dir, `source.${inExt}`)
    fs.writeFileSync(inputPath, input)
    await execFileAsync(soffice, ['--headless', '--convert-to', convertTo, '--outdir', dir, inputPath], {
      timeout: 60000
    })
    return fs.readFileSync(path.join(dir, `source.${outExt}`))
  } finally {
    fs.rmSync(dir, { recursive: true, force: true })
  }
}

/**
 * Retries a `soffice` conversion a few times before giving up: when several
 * documents are converted back-to-back (the "Download all" bundle converts 3
 * PDFs in a row, or docxToPageImages converts one PDF page at a time), a
 * prior conversion's `soffice` instance can still be shutting down and
 * holding the user-profile lock, so the next call fails transiently. A short
 * wait-and-retry lets that instance exit — without it, later conversions in
 * a batch silently dropped out.
 */
async function convertWithRetry(
  soffice: string,
  input: Buffer,
  inExt: string,
  convertTo: string,
  outExt: string = convertTo
): Promise<Buffer> {
  let lastErr: unknown
  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      return await sofficeConvert(soffice, input, inExt, convertTo, outExt)
    } catch (err) {
      lastErr = err
      await new Promise((resolve) => setTimeout(resolve, 700 * (attempt + 1)))
    }
  }
  throw new Error(
    'PDF conversion requires LibreOffice to be installed on this computer. ' +
      'Install it from libreoffice.org, then try again — or use the Word (.docx) or Print option instead.\n' +
      (lastErr instanceof Error ? lastErr.message : String(lastErr))
  )
}

/** Convert a .docx buffer to PDF via a local LibreOffice install. Throws a clear error when LibreOffice isn't installed, instead of a cryptic ENOENT. */
export async function convertDocxToPdf(docxBuffer: Buffer): Promise<Buffer> {
  return convertWithRetry(findSofficeBinary(), docxBuffer, 'docx', 'pdf')
}

/**
 * Renders a .docx as one PNG per page — LibreOffice's own docx→PDF
 * conversion (convertDocxToPdf), then LibreOffice's own PDF→PNG raster
 * export for each page in turn (pdf-lib splits the PDF into single-page
 * PDFs first, since `soffice --convert-to png` only ever rasterizes a
 * document's first page).
 *
 * This is the app's accurate document preview/print path — deliberately
 * NOT pdf.js (see src/pdfToImages.ts, used elsewhere for scanned-PDF OCR):
 * a LibreOffice-produced PDF for a font the machine doesn't have installed
 * (e.g. "Book Antiqua", "Segoe UI" — both Microsoft-only, absent from every
 * non-Windows LibreOffice install) embeds a substituted font whose glyph
 * mapping pdf.js decodes wrong, silently swapping in different characters
 * (reported as: "Quthbullapur" rendering as "8uthbulla/ur") — confirmed by
 * rendering the exact same, byte-verified-correct PDF through Poppler
 * (correct) and pdf.js (garbled) side by side. LibreOffice's own rasterizer
 * (used here) renders that same substituted font correctly, because it's
 * the same engine that chose the substitution in the first place.
 */
// LibreOffice's plain `--convert-to png` rasterizes at 96dpi (screen
// resolution) — fine on screen, since PAGE_WIDTH/PAGE_HEIGHT are that same
// 96dpi A4 size and the browser shows the PNG 1:1, but that image is also
// what gets sent to the printer/PDF export, where 96dpi stretched across a
// real page comes out visibly soft — "photo print" rather than crisp text.
// Requesting PixelWidth/PixelHeight explicitly (computed per page from the
// PDF's own point size, so it holds for any page size/orientation, not just
// portrait A4) renders at a real print resolution instead.
const PRINT_DPI = 300

export async function docxToPageImages(docxBuffer: Buffer): Promise<Buffer[]> {
  const soffice = findSofficeBinary()
  const pdf = await convertWithRetry(soffice, docxBuffer, 'docx', 'pdf')
  const src = await PDFDocument.load(pdf)
  const pageCount = src.getPageCount()

  const images: Buffer[] = new Array(pageCount)
  for (let i = 0; i < pageCount; i++) {
    const single = await PDFDocument.create()
    const [copied] = await single.copyPages(src, [i])
    single.addPage(copied)
    const { width, height } = copied.getSize()
    const pixelWidth = Math.round((width / 72) * PRINT_DPI)
    const pixelHeight = Math.round((height / 72) * PRINT_DPI)
    const filter = `png:draw_png_Export:{"PixelWidth":{"type":"long","value":"${pixelWidth}"},"PixelHeight":{"type":"long","value":"${pixelHeight}"}}`
    images[i] = await convertWithRetry(soffice, Buffer.from(await single.save()), 'pdf', filter, 'png')
  }
  return images
}

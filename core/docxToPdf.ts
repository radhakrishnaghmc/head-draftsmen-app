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
  if (!found) {
    // Thrown straight through the IPC bridge to the renderer with no
    // wrapping (unlike convertWithRetry's own message below, which only
    // covers a binary that exists but fails to run) — so this needs to be
    // the actionable, user-facing text itself, not a diagnostic stub.
    throw new Error(
      'PDF conversion requires LibreOffice to be installed on this computer. ' +
        'Install it from libreoffice.org, then try again — or use the Word (.docx) or Print option instead.'
    )
  }
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

/**
 * Same idea as sofficeConvert, but N input files converted in ONE soffice
 * invocation instead of N separate ones. `soffice`'s own process-startup
 * cost (loading config/fonts) dominates a single conversion's wall time far
 * more than the actual rendering work does — this matters most on Windows,
 * where that startup cost is markedly higher than on macOS/Linux — so
 * passing every file to one invocation pays that cost once instead of once
 * per file. Measured on a real 5-page document (page->PNG step only): 11.0s
 * as 5 separate calls vs 3.85s batched (2.9x) — byte-identical output — and
 * the win scales with page count. `names` must be unique (no extension) —
 * they become both the input and output filename stem, so the caller can
 * map results back to the page each one came from.
 */
async function sofficeConvertBatch(
  soffice: string,
  inputs: { name: string; buffer: Buffer }[],
  inExt: string,
  convertTo: string,
  outExt: string = convertTo
): Promise<Map<string, Buffer>> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'docugen-convert-'))
  try {
    const inputPaths = inputs.map(({ name, buffer }) => {
      const p = path.join(dir, `${name}.${inExt}`)
      fs.writeFileSync(p, buffer)
      return p
    })
    await execFileAsync(soffice, ['--headless', '--convert-to', convertTo, '--outdir', dir, ...inputPaths], {
      // Scales with batch size — a single file's 60s budget wouldn't be
      // enough headroom for a many-page document processed in one invocation.
      timeout: 60000 + inputs.length * 5000
    })
    const results = new Map<string, Buffer>()
    for (const { name } of inputs) {
      results.set(name, fs.readFileSync(path.join(dir, `${name}.${outExt}`)))
    }
    return results
  } finally {
    fs.rmSync(dir, { recursive: true, force: true })
  }
}

/** Batched counterpart to convertWithRetry — same transient-lock retry, for sofficeConvertBatch. */
async function convertWithRetryBatch(
  soffice: string,
  inputs: { name: string; buffer: Buffer }[],
  inExt: string,
  convertTo: string,
  outExt: string = convertTo
): Promise<Map<string, Buffer>> {
  let lastErr: unknown
  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      return await sofficeConvertBatch(soffice, inputs, inExt, convertTo, outExt)
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

  // Split into single-page PDFs first (soffice's own --convert-to only ever
  // rasterizes a document's first page), then group by pixel size — the
  // PixelWidth/PixelHeight filter applies to a whole soffice invocation, not
  // per file, so pages that differ in size (rare, but not impossible) need
  // their own batch. In practice this is almost always one group covering
  // every page, which is exactly what makes batching such a large win: see
  // sofficeConvertBatch's own doc comment for the measured 2.9x.
  const groups = new Map<string, { name: string; buffer: Buffer }[]>()
  const pageIndexByName = new Map<string, number>()
  for (let i = 0; i < pageCount; i++) {
    const single = await PDFDocument.create()
    const [copied] = await single.copyPages(src, [i])
    single.addPage(copied)
    const { width, height } = copied.getSize()
    const pixelWidth = Math.round((width / 72) * PRINT_DPI)
    const pixelHeight = Math.round((height / 72) * PRINT_DPI)
    const key = `${pixelWidth}x${pixelHeight}`
    const name = `page-${i}`
    pageIndexByName.set(name, i)
    const bucket = groups.get(key) ?? []
    bucket.push({ name, buffer: Buffer.from(await single.save()) })
    groups.set(key, bucket)
  }

  const images: Buffer[] = new Array(pageCount)
  for (const [key, inputs] of groups) {
    const [pixelWidth, pixelHeight] = key.split('x')
    const filter = `png:draw_png_Export:{"PixelWidth":{"type":"long","value":"${pixelWidth}"},"PixelHeight":{"type":"long","value":"${pixelHeight}"}}`
    const results = await convertWithRetryBatch(soffice, inputs, 'pdf', filter, 'png')
    for (const [name, buffer] of results) {
      images[pageIndexByName.get(name)!] = buffer
    }
  }
  return images
}

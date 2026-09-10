import { execFile } from 'child_process'
import { promisify } from 'util'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { PDFDocument } from 'pdf-lib'
import PizZip from 'pizzip'

const execFileAsync = promisify(execFile)

/**
 * The header2 letterhead templates (and a couple of others) declare their
 * Telugu-script caption run as "Gautami" (see word2007Compat.ts's doc
 * comment on stripEmbeddedFonts for why: a prior fix dropped the embedded
 * Noto Sans Telugu font in favor of naming a Windows-native complex-script
 * font, since Word 2007 can't open a document with the newer OOXML
 * font-embedding elements). Real Word on a government office's Windows PC
 * has Gautami built in and renders it fine.
 *
 * LibreOffice, though, doesn't ship Gautami and its own font-substitution
 * logic doesn't pick a complex-script-capable stand-in for it — it reaches
 * for "Arial Unicode MS" first, which has Telugu code points mapped to
 * glyphs but no OpenType shaping rules for them, so conjuncts/matras render
 * in the wrong shapes and positions (confirmed by rendering the exact same
 * template through LibreOffice with and without this substitution, side by
 * side: "Gautami" comes out visibly garbled, "Noto Sans Telugu" doesn't).
 * This bites every LibreOffice-based PDF export (this file's whole reason
 * to exist), regardless of which OS it runs on.
 *
 * The fix: rename just the font declaration, only in the buffer handed to
 * `soffice` for conversion — never the .docx the user actually keeps, which
 * must stay declaring Gautami for Word. No font needs bundling for this:
 * LibreOffice ships Noto Sans Telugu inside its own app bundle on every
 * platform (Mac/Windows/Linux all include it under its install dir's own
 * `fonts` folder), so it's always available to `soffice` even on a machine
 * with no Telugu font installed at the OS level.
 */
const PDF_FONT_SUBSTITUTES: Record<string, string> = {
  Gautami: 'Noto Sans Telugu'
}

const FONT_ATTR_NAMES = ['ascii', 'hAnsi', 'cs', 'eastAsia']

/** Exported for testing — see PDF_FONT_SUBSTITUTES above for why this exists. */
export function substituteFontsForLibreOffice(docxBuffer: Buffer): Buffer {
  const zip = new PizZip(docxBuffer)
  const partNames = Object.keys(zip.files).filter(
    (name) => name === 'word/document.xml' || name === 'word/styles.xml' || /^word\/(header|footer)\d*\.xml$/.test(name)
  )
  let anyChange = false
  for (const name of partNames) {
    const file = zip.file(name)
    if (!file) continue
    let xml = file.asText()
    let fileChanged = false
    for (const [from, to] of Object.entries(PDF_FONT_SUBSTITUTES)) {
      const re = new RegExp(`(w:(?:${FONT_ATTR_NAMES.join('|')})=")${from}(")`, 'g')
      if (re.test(xml)) {
        xml = xml.replace(re, `$1${to}$2`)
        fileChanged = true
      }
    }
    if (fileChanged) {
      zip.file(name, xml)
      anyChange = true
    }
  }
  return anyChange ? Buffer.from(zip.generate({ type: 'nodebuffer' })) : docxBuffer
}

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
    // Only the temp file soffice actually reads gets the font rename (see
    // substituteFontsForLibreOffice above) — `input` itself is left alone,
    // so a caller that also writes the same buffer out as a .docx (e.g. the
    // "Download all" bundle) still gets the Word-2007-safe Gautami original.
    fs.writeFileSync(inputPath, inExt === 'docx' ? substituteFontsForLibreOffice(input) : input)
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
 * conversion (convertDocxToPdf), then rasterized to PNG by the
 * `rasterizePages` callback the caller injects.
 *
 * This file must never import `electron` (it also runs in contexts that
 * don't have it), so the actual PDF→PNG rasterization — which does need a
 * real browser engine — lives in electron/pdfRaster.ts and is passed in
 * here instead of called directly. See that file's doc comment for why:
 * in short, LibreOffice's own `--convert-to png` rasterizer turned out to
 * have its own, separate bug mis-shaping Telugu/complex-script glyphs, so
 * it's no longer used for this path.
 *
 * Deliberately not pdf.js either (see src/pdfToImages.ts, used elsewhere for
 * scanned-PDF OCR): a LibreOffice-produced PDF for a font the machine
 * doesn't have installed (e.g. "Book Antiqua", "Segoe UI" — both
 * Microsoft-only, absent from every non-Windows LibreOffice install) embeds
 * a substituted font whose glyph mapping pdf.js decodes wrong, silently
 * swapping in different characters (reported as: "Quthbullapur" rendering
 * as "8uthbulla/ur") — confirmed by rendering the exact same,
 * byte-verified-correct PDF through Poppler (correct) and pdf.js (garbled)
 * side by side. Electron's bundled Chromium/PDFium PDF viewer renders both
 * that substituted font AND Telugu correctly, which is why it's used here.
 */
// Requesting pixel dimensions computed from the PDF's own point size (rather
// than rasterizing at screen resolution) keeps the printed/exported page
// crisp instead of "photo print" soft. 300dpi is real print resolution.
const PRINT_DPI = 300

export async function docxToPageImages(
  docxBuffer: Buffer,
  rasterizePages: (pdf: Buffer, pageSizes: { pixelWidth: number; pixelHeight: number }[]) => Promise<Buffer[]>
): Promise<Buffer[]> {
  const soffice = findSofficeBinary()
  const pdf = await convertWithRetry(soffice, docxBuffer, 'docx', 'pdf')
  const src = await PDFDocument.load(pdf)
  const pageSizes = src.getPages().map((page) => {
    const { width, height } = page.getSize()
    return {
      pixelWidth: Math.round((width / 72) * PRINT_DPI),
      pixelHeight: Math.round((height / 72) * PRINT_DPI)
    }
  })
  return rasterizePages(pdf, pageSizes)
}

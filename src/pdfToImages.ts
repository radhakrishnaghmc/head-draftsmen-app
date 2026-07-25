import * as pdfjsLib from 'pdfjs-dist'
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url'

// Not yet in this project's configured TS lib version's typings.
declare global {
  interface Uint8Array {
    toHex(): string
  }
}

// pdfjs-dist assumes the brand-new (TC39 "Uint8Array to/from base64/hex")
// Uint8Array.prototype.toHex() exists, unconditionally, calling it while
// hashing a PDF's content during load — but the Chromium build bundled with
// this app's current Electron version predates that method entirely, so
// every PDF load throws "toHex is not a function" with no fallback path.
// Confirmed as this exact, known pdf.js/runtime compatibility gap (not a
// bug in a specific PDF) — polyfilled here per the TC39 spec (each byte as
// 2-digit lowercase hex, concatenated) rather than pinning to an older,
// less-maintained pdfjs-dist release.
if (typeof Uint8Array.prototype.toHex !== 'function') {
  Object.defineProperty(Uint8Array.prototype, 'toHex', {
    value(this: Uint8Array) {
      let hex = ''
      for (const byte of this) hex += byte.toString(16).padStart(2, '0')
      return hex
    },
    writable: true,
    configurable: true
  })
}

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl

// Scale 2 off a PDF's standard 72dpi page is ~144dpi — plenty of detail for
// OCR without producing an oversized image; electron/ocr.ts caps input size
// further before running OCR regardless.
const RENDER_SCALE = 2

/**
 * Renders every page of a PDF (a scanned/photographed multi-page estimate
 * saved as one PDF file, rather than separate photo files) to a PNG data
 * URL, one per page, in page order — so a PDF upload can feed the exact same
 * per-photo OCR pipeline (electron/ocr.ts's recognizeImage via the
 * ocrEstimatePhotos IPC call) that photo uploads already use, with no
 * changes needed downstream.
 */
export async function pdfPagesToDataUrls(file: File): Promise<string[]> {
  const buffer = await file.arrayBuffer()
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise
  const dataUrls: string[] = []
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum)
    const viewport = page.getViewport({ scale: RENDER_SCALE })
    const canvas = document.createElement('canvas')
    canvas.width = viewport.width
    canvas.height = viewport.height
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Could not get a 2D canvas context to render the PDF.')
    await page.render({ canvas, canvasContext: context, viewport }).promise
    dataUrls.push(canvas.toDataURL('image/png'))
  }
  return dataUrls
}

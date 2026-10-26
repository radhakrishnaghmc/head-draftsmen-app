import { PDFDocument } from 'pdf-lib'

// Requesting pixel dimensions computed from the PDF's own point size (rather
// than rasterizing at screen resolution) keeps the printed/exported page
// crisp instead of "photo print" soft. 300dpi is real print resolution.
const PRINT_DPI = 300

/**
 * Renders a .docx as one PNG per page: `toPdf` converts it to a PDF, then
 * `rasterizePages` turns each page into a PNG at an exact pixel size.
 *
 * This file must never import `electron` (it also runs in contexts that
 * don't have it), so both steps — which need a real browser engine — live in
 * electron/ (docxChromiumPdf.ts and pdfRaster.ts) and are passed in.
 *
 * Deliberately not pdf.js for the rasterizing (see src/pdfToImages.ts, used
 * elsewhere for scanned-PDF OCR): Electron's bundled Chromium/PDFium viewer
 * renders substituted and Telugu fonts correctly where pdf.js can decode a
 * substituted font's glyph mapping wrong.
 */
export async function docxToPageImages(
  docxBuffer: Buffer,
  toPdf: (docx: Buffer) => Promise<Buffer>,
  rasterizePages: (pdf: Buffer, pageSizes: { pixelWidth: number; pixelHeight: number }[]) => Promise<Buffer[]>
): Promise<Buffer[]> {
  const pdf = await toPdf(docxBuffer)
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

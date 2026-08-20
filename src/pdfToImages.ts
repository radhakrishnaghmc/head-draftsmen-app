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
  return pdfPagesToDataUrlsFromData(await file.arrayBuffer())
}

/**
 * Same as pdfPagesToDataUrls, from raw bytes — used for a PDF produced
 * in-process (e.g. LibreOffice's docx→PDF conversion, see core/docxToPdf.ts)
 * rather than read from a file input. `scale` overrides RENDER_SCALE — the
 * accurate document-preview renderer (see docPage.ts's renderAccurate) wants
 * a sharper image than the OCR pipeline needs.
 */
export async function pdfPagesToDataUrlsFromData(
  data: ArrayBuffer | Uint8Array,
  scale: number = RENDER_SCALE
): Promise<string[]> {
  // Loaded on demand — see pdfToText.ts.
  const { pdfjsLib } = await import('./pdfjsSetup')
  const pdf = await pdfjsLib.getDocument({ data }).promise
  const dataUrls: string[] = []
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum)
    const viewport = page.getViewport({ scale })
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

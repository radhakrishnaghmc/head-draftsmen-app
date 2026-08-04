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
  // Loaded on demand — see pdfToText.ts.
  const { pdfjsLib } = await import('./pdfjsSetup')
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

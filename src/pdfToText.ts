import { SELECTED_ROW_MARK } from '../core/selectedRowMark'

interface TextItem {
  str: string
  transform: number[]
  width?: number
}


/**
 * When two or more bidders tie at L-1, the portal's price-bid table still
 * shows every one as "L-1" and ticks the "Select" radio of the one the
 * department picked — a purely graphical cue the text layer doesn't carry.
 * Renders the page and looks for the blue (filled) radio under the "Select"
 * header at each L-1 row; returns the y (PDF space) of the single ticked row,
 * or undefined when the page has no tie or the tick can't be told apart.
 */
async function findSelectedTieRowY(page: any, items: TextItem[]): Promise<number | undefined> {
  const l1 = items.filter((i) => i.str.trim() === 'L-1')
  if (l1.length < 2) return undefined
  const header = items.find((i) => i.str.trim() === 'Select')
  if (!header) return undefined
  const viewport = page.getViewport({ scale: 1 })
  const canvas = document.createElement('canvas')
  canvas.width = viewport.width
  canvas.height = viewport.height
  const ctx = canvas.getContext('2d', { willReadFrequently: true })
  if (!ctx) return undefined
  await page.render({ canvas, canvasContext: ctx, viewport }).promise
  const cx = header.transform[4] + (header.width ?? 0) / 2
  let best: { y: number; blue: number } | undefined
  let hits = 0
  for (const it of l1) {
    const y = it.transform[5]
    const [px, py] = viewport.convertToViewportPoint(cx, y + 4)
    const img = ctx.getImageData(Math.max(0, Math.round(px) - 8), Math.max(0, Math.round(py) - 8), 17, 17).data
    let blue = 0
    for (let k = 0; k < img.length; k += 4) if (img[k + 2] > 200 && img[k] < 90 && img[k + 1] > 70 && img[k + 1] < 190) blue++
    if (blue >= 6) {
      hits++
      if (!best || blue > best.blue) best = { y, blue }
    }
  }
  return hits === 1 ? best!.y : undefined
}

/**
 * Extracts a digital PDF's selectable text as reconstructed lines, in
 * reading order — text items are grouped by their rounded y-position into
 * lines, and sorted left-to-right within each line, so a portal report's
 * table rows come back as whole lines (e.g. "<Company> <ECV> Less <pct>
 * <amount> L-1") that core parsers can read. Meant for the digital,
 * text-bearing PDFs the tender portal produces — not scanned image PDFs,
 * which have no selectable text and would need OCR (see src/pdfToImages.ts).
 */
export async function pdfToTextLines(file: File): Promise<string[]> {
  return pdfToTextLinesFromData(await file.arrayBuffer())
}

/** Same as pdfToTextLines, from raw bytes — used for folder files read (as base64) via the main process. */
export async function pdfToTextLinesFromData(data: ArrayBuffer | Uint8Array): Promise<string[]> {
  return (await pdfPagesToTextLinesFromData(data)).flat()
}

/**
 * Like pdfToTextLines, but falls back to OCR when the PDF carries no
 * selectable text layer at all — e.g. an eGP portal "Stage Selected Form" /
 * L1 sheet saved via Microsoft Print to PDF from a canvas-rendered popup,
 * which rasterizes the whole page instead of keeping real text (pdf.js then
 * returns zero text items rather than an error, so the difference has to be
 * detected here, not caught as an exception). Renders each page to an image
 * (see pdfToImages.ts) and OCRs it through the same pipeline photo uploads
 * use — every L1 sheet / tender-evaluation upload should go through this
 * instead of pdfToTextLines directly, so this PDF variant reads like any
 * other.
 */
export async function pdfToTextLinesOrOcr(file: File): Promise<string[]> {
  const lines = await pdfToTextLines(file)
  if (lines.length > 0) return lines
  const [{ pdfPagesToDataUrlsFromData }, { api }, { gluedLineCount }] = await Promise.all([
    import('./pdfToImages'),
    import('./ipc'),
    import('../core/ocrSpacing')
  ])
  const data = await file.arrayBuffer()
  // The recognizer drops word spaces at some render scales (see
  // core/ocrSpacing.ts) — start at the default and only re-read at other
  // scales when the result shows glued words, keeping the cleanest read.
  let best: string[] = []
  let bestGlued = Infinity
  for (const scale of [2, 2.8, 1.4]) {
    const result = await api.ocrPhotosToLines(await pdfPagesToDataUrlsFromData(data, scale))
    const glued = gluedLineCount(result)
    if (glued < bestGlued) {
      best = result
      bestGlued = glued
    }
    if (glued === 0) break
  }
  return best
}

/**
 * Same reconstruction as pdfToTextLinesFromData, but keeping each page's
 * lines separate — lets a caller tell whether a *specific* page has a real
 * text layer (e.g. to decide OCR vs. direct extraction per page of a mixed
 * scanned/digital PDF) rather than only the document as a whole.
 */
export async function pdfPagesToTextLinesFromData(
  data: ArrayBuffer | Uint8Array,
  onProgress?: (page: number, totalPages: number) => void
): Promise<string[][]> {
  // Loaded on demand — pdfjs (and its worker) is heavy and only needed when the
  // user actually reads a PDF, so it stays out of the initial startup bundle.
  const { pdfjsLib } = await import('./pdfjsSetup')
  const loadingTask = pdfjsLib.getDocument({ data })
  const pdf = await loadingTask.promise
  try {
    const pages: string[][] = []
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum)
      const content = await page.getTextContent()
      const byRow = new Map<number, { x: number; s: string }[]>()
      for (const item of content.items as TextItem[]) {
        if (!item.str) continue
        const y = Math.round(item.transform[5])
        if (!byRow.has(y)) byRow.set(y, [])
        byRow.get(y)!.push({ x: item.transform[4], s: item.str })
      }
      // Tied L-1s: tag the row whose Select radio is ticked (best-effort — a
      // render failure just leaves the page untagged, i.e. the old first-L-1 read).
      let selectedY: number | undefined
      try {
        const sy = await findSelectedTieRowY(page, content.items as TextItem[])
        if (sy != null) selectedY = Math.round(sy)
      } catch {
        selectedY = undefined
      }
      if (selectedY != null && byRow.has(selectedY)) byRow.get(selectedY)!.push({ x: Infinity, s: SELECTED_ROW_MARK })
      const lines: string[] = []
      for (const y of [...byRow.keys()].sort((a, b) => b - a)) {
        const line = byRow
          .get(y)!
          .sort((a, b) => a.x - b.x)
          .map((c) => c.s)
          .join(' ')
          .replace(/\s+/g, ' ')
          .trim()
        if (line) lines.push(line)
      }
      pages.push(lines)
      onProgress?.(pageNum, pdf.numPages)
    }
    return pages
  } finally {
    // A folder scan opens dozens/hundreds of these in a tight loop — without
    // destroying each document, pdf.js keeps its worker-side font/image
    // caches alive per document and the renderer eventually OOMs partway
    // through a large batch. destroy() lives on the loading task, not the
    // resolved PDFDocumentProxy.
    await loadingTask.destroy()
  }
}

/** A reconstructed line plus its position: `x` is the line's leftmost item, `y`
 * its (page-offset) vertical position in reading order. Lets a caller tell table
 * columns apart by x — e.g. the company-name column vs the comments column on
 * the "Bidders Made Non-Responsive" sheet — which the plain text loses. */
export interface PositionedLine {
  text: string
  x: number
  y: number
}

/** Like pdfToTextLines, but each line keeps its leftmost x and its y (offset per
 * page so cross-page ys stay in reading order and never collide). */
export async function pdfToPositionedLines(file: File): Promise<PositionedLine[]> {
  const { pdfjsLib } = await import('./pdfjsSetup')
  const loadingTask = pdfjsLib.getDocument({ data: await file.arrayBuffer() })
  const pdf = await loadingTask.promise
  try {
    const out: PositionedLine[] = []
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum)
      const content = await page.getTextContent()
      const pageOffset = (pageNum - 1) * 100000 // keep later pages strictly below earlier ones
      const byRow = new Map<number, { x: number; s: string }[]>()
      for (const item of content.items as TextItem[]) {
        if (!item.str) continue
        const y = Math.round(item.transform[5])
        if (!byRow.has(y)) byRow.set(y, [])
        byRow.get(y)!.push({ x: item.transform[4], s: item.str })
      }
      for (const y of [...byRow.keys()].sort((a, b) => b - a)) {
        const cells = byRow.get(y)!.sort((a, b) => a.x - b.x)
        const text = cells
          .map((c) => c.s)
          .join(' ')
          .replace(/\s+/g, ' ')
          .trim()
        if (text) out.push({ text, x: cells[0].x, y: pageOffset + (100000 - y) })
      }
    }
    return out
  } finally {
    await loadingTask.destroy()
  }
}

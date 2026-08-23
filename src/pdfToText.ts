interface TextItem {
  str: string
  transform: number[]
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
 * Same reconstruction as pdfToTextLinesFromData, but keeping each page's
 * lines separate — lets a caller tell whether a *specific* page has a real
 * text layer (e.g. to decide OCR vs. direct extraction per page of a mixed
 * scanned/digital PDF) rather than only the document as a whole.
 */
export async function pdfPagesToTextLinesFromData(data: ArrayBuffer | Uint8Array): Promise<string[][]> {
  // Loaded on demand — pdfjs (and its worker) is heavy and only needed when the
  // user actually reads a PDF, so it stays out of the initial startup bundle.
  const { pdfjsLib } = await import('./pdfjsSetup')
  const pdf = await pdfjsLib.getDocument({ data }).promise
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
  }
  return pages
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
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise
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
}

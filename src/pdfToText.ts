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
  // Loaded on demand — pdfjs (and its worker) is heavy and only needed when the
  // user actually reads a PDF, so it stays out of the initial startup bundle.
  const { pdfjsLib } = await import('./pdfjsSetup')
  const pdf = await pdfjsLib.getDocument({ data }).promise
  const lines: string[] = []
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
  }
  return lines
}

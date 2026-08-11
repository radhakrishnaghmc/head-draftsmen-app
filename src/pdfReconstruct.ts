import type { PageGeometry, GeomText, GeomHLine, GeomVLine } from '@core/reconstructDoc'

/**
 * Extract, per page, the geometry a PDF needs to be rebuilt as an editable Word
 * document that preserves its layout: the positioned text (with bold), and the
 * ruling lines (table borders) recovered from the PDF's drawing operators.
 *
 * Only meaningful for text-based PDFs (a real, selectable-text document); a
 * scanned/image PDF yields no text and the caller falls back to OCR / the
 * LibreOffice image conversion. All the layout maths (grid detection, cell
 * assignment) lives in the pure core/reconstructDoc — this only harvests the raw
 * geometry from pdf.js.
 */
export async function extractPdfGeometry(data: ArrayBuffer | Uint8Array): Promise<PageGeometry[]> {
  const { pdfjsLib } = await import('./pdfjsSetup')
  const pdf = await pdfjsLib.getDocument({ data }).promise
  const pages: PageGeometry[] = []
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum)
    const viewport = page.getViewport({ scale: 1 })
    const width = viewport.width
    const height = viewport.height

    // Ruling lines from the draw operators. Each constructPath carries a minMax
    // bounding box [minX,minY,maxX,maxY] (PDF user space, y up); a thin-wide box
    // is a horizontal rule, a tall-thin box a vertical rule, and a large box a
    // rectangle (its four sides). y is flipped to top-down to match the text.
    const ops = await page.getOperatorList()
    const constructPath = pdfjsLib.OPS.constructPath
    const hlines: GeomHLine[] = []
    const vlines: GeomVLine[] = []
    for (let i = 0; i < ops.fnArray.length; i++) {
      if (ops.fnArray[i] !== constructPath) continue
      const mm = ops.argsArray[i]?.[2] as ArrayLike<number> | undefined
      if (!mm || mm.length < 4) continue
      const minX = mm[0]
      const minY = mm[1]
      const maxX = mm[2]
      const maxY = mm[3]
      const w = maxX - minX
      const h = maxY - minY
      const yTop = height - maxY
      const yBot = height - minY
      if (h < 3 && w > 15) {
        hlines.push({ y: (yTop + yBot) / 2, x1: minX, x2: maxX })
      } else if (w < 3 && h > 15) {
        vlines.push({ x: (minX + maxX) / 2, y1: yTop, y2: yBot })
      } else if (w > 15 && h > 15) {
        hlines.push({ y: yTop, x1: minX, x2: maxX })
        hlines.push({ y: yBot, x1: minX, x2: maxX })
        vlines.push({ x: minX, y1: yTop, y2: yBot })
        vlines.push({ x: maxX, y1: yTop, y2: yBot })
      }
    }

    // Positioned text. item.fontName is an internal id ("g_d0_f1"); the real
    // font name (…Bold…) comes from the resolved font object, cached per page.
    const content = await page.getTextContent()
    const boldByFont = new Map<string, boolean>()
    const isBold = (fontName: string): boolean => {
      if (!fontName) return false
      const cached = boldByFont.get(fontName)
      if (cached !== undefined) return cached
      let bold = false
      try {
        const font = page.commonObjs.get(fontName) as { name?: string } | undefined
        bold = !!font?.name && /bold|black|heavy|semibold/i.test(font.name)
      } catch {
        bold = false
      }
      boldByFont.set(fontName, bold)
      return bold
    }
    const texts: GeomText[] = []
    for (const item of content.items as { str: string; transform: number[]; width: number; fontName: string }[]) {
      if (!item.str || !item.str.trim()) continue
      const t = item.transform
      // Font size in points ≈ the vertical scale of the text transform.
      const size = Math.hypot(t[2], t[3]) || Math.abs(t[3]) || 11
      texts.push({
        str: item.str,
        x: t[4],
        y: height - t[5],
        w: item.width,
        size,
        bold: isBold(item.fontName)
      })
    }

    pages.push({ width, height, texts, hlines, vlines })
  }
  return pages
}

/**
 * Rebuild a Word (HTML → .docx) document that reproduces a text-based PDF's
 * layout accurately: real editable tables (from the PDF's own ruling lines),
 * paragraphs, bold, and centering — not the flat OCR word stream that loses all
 * structure. The renderer extracts the geometry with pdf.js
 * (src/pdfReconstruct.ts); this module is the pure, testable transform from that
 * geometry to HTML, handed to core/htmlToDocx's convertHtmlToDocx in the main
 * process.
 */

/** One piece of text with its top-left position (page coords, y measured from the top) and whether it's bold. */
export interface GeomText {
  str: string
  x: number
  y: number
  /** Advance width of the text run, so a line's right edge (and thus its centre) is known. */
  w: number
  /** Font size in points, so the Word output keeps the PDF's real text sizes. */
  size: number
  bold: boolean
}
/** A horizontal ruling line at `y`, spanning x1..x2. */
export interface GeomHLine {
  y: number
  x1: number
  x2: number
}
/** A vertical ruling line at `x`, spanning y1..y2. */
export interface GeomVLine {
  x: number
  y1: number
  y2: number
}
export interface PageGeometry {
  width: number
  height: number
  texts: GeomText[]
  hlines: GeomHLine[]
  vlines: GeomVLine[]
}

import type { DocBlock, DocParagraph, DocRun } from './docxBuilder'

// 1 pt = 20 twips. The letter text column (12240 − 2×1440 margins) is 9360 twips;
// a table wider than that is scaled down to fit, otherwise real widths are kept.
const TWIPS_PER_PT = 20
const TEXT_WIDTH_TWIPS = 9360

/** A run carrying the PDF's real font size (half-points) so the Word output keeps the text sizes. */
function run(text: string, bold: boolean, sizePt: number): DocRun {
  return { text, bold, size: Math.max(2, Math.round(sizePt * 2)) }
}

/** Collapse near-equal numbers (within `tol`) into single averaged, sorted values — used to turn many overlapping ruling-line coordinates into clean row/column boundaries. */
function clusterVals(vals: number[], tol: number): number[] {
  const sorted = [...vals].sort((a, b) => a - b)
  const groups: { c: number; sum: number; n: number }[] = []
  for (const v of sorted) {
    const last = groups[groups.length - 1]
    if (last && Math.abs(v - last.c) <= tol) {
      last.sum += v
      last.n += 1
      last.c = last.sum / last.n
    } else {
      groups.push({ c: v, sum: v, n: 1 })
    }
  }
  return groups.map((g) => Math.round(g.c))
}

interface Line {
  /** One run per line (bold applies to the whole line). */
  text: string
  bold: boolean
  /** Font size in points (median of the line's pieces). */
  size: number
  /** Horizontal centre of the line (for centre-alignment detection). */
  cx: number
}

/** Group a set of text pieces into visual lines (by y), each read left-to-right. */
function linesOf(texts: GeomText[]): Line[] {
  if (texts.length === 0) return []
  const rowYs = clusterVals(
    texts.map((t) => Math.round(t.y)),
    6
  )
  const lines: Line[] = []
  for (const ry of rowYs) {
    const row = texts.filter((t) => Math.abs(t.y - ry) <= 6).sort((a, b) => a.x - b.x)
    if (row.length === 0) continue
    const text = row
      .map((t) => t.str)
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim()
    if (!text) continue
    const left = row[0].x
    const last = row[row.length - 1]
    const sizes = row.map((t) => t.size).sort((a, b) => a - b)
    lines.push({
      text,
      bold: row.some((t) => t.bold),
      size: sizes[Math.floor(sizes.length / 2)] || 11,
      cx: (left + last.x + last.w) / 2
    })
  }
  return lines
}

/** Free text (outside any table) as paragraphs, centred when the line sits near the page centre. */
function paragraphsBlocks(texts: GeomText[], pageWidth: number): DocParagraph[] {
  return linesOf(texts).map((l) => ({
    kind: 'paragraph',
    runs: [run(l.text, l.bold, l.size)],
    align: Math.abs(l.cx - pageWidth / 2) < pageWidth * 0.14 ? 'center' : 'left'
  }))
}

/**
 * Add row boundaries inside tall ruled bands that actually hold several logical
 * rows with no rules between them (common in forms: a big label/value box). A
 * band is subdivided — at the mid-points between its text lines — only when TWO
 * OR MORE columns have multiple parallel text lines; a band where just one column
 * wraps (e.g. a long value in an otherwise single-row cell) is left intact.
 */
function refineRowYs(baseRowYs: number[], colXs: number[], tableTexts: GeomText[]): number[] {
  const out: number[] = []
  for (let r = 0; r < baseRowYs.length - 1; r++) {
    const ry0 = baseRowYs[r]
    const ry1 = baseRowYs[r + 1]
    out.push(ry0)
    const band = tableTexts.filter((t) => t.y >= ry0 - 3 && t.y < ry1)
    const perColLineYs: number[][] = []
    let colsWithMulti = 0
    for (let c = 0; c < colXs.length - 1; c++) {
      const colTexts = band.filter((t) => t.x + 2 >= colXs[c] && t.x < colXs[c + 1])
      const lineYs = clusterVals(
        colTexts.map((t) => Math.round(t.y)),
        6
      )
      perColLineYs.push(lineYs)
      if (lineYs.length >= 2) colsWithMulti++
    }
    if (colsWithMulti >= 2) {
      const allYs = clusterVals(perColLineYs.flat(), 6)
      for (let i = 0; i < allYs.length - 1; i++) out.push((allYs[i] + allYs[i + 1]) / 2)
    }
  }
  out.push(baseRowYs[baseRowYs.length - 1])
  return out
}

/** Build the doc-model blocks for one page — its table (from ruling lines) plus the paragraphs above/below. */
function pageBlocks(g: PageGeometry): DocBlock[] {
  const rowYs = clusterVals(
    g.hlines.filter((l) => l.y > 5 && l.y < g.height - 5).map((l) => l.y),
    3
  )
  const colXs = clusterVals(
    g.vlines.filter((l) => l.x > 5 && l.x < g.width - 5).map((l) => l.x),
    3
  )

  // No usable grid → the whole page is just paragraphs.
  if (rowYs.length < 2 || colXs.length < 2) return paragraphsBlocks(g.texts, g.width)

  const yTop = rowYs[0]
  const yBot = rowYs[rowYs.length - 1]
  const xL = colXs[0]
  const xR = colXs[colXs.length - 1]
  const inTable = (t: GeomText) => t.y >= yTop - 3 && t.y <= yBot + 3 && t.x >= xL - 3 && t.x <= xR + 3
  const tableTexts = g.texts.filter(inTable)
  const above = g.texts.filter((t) => t.y < yTop - 3)
  const below = g.texts.filter((t) => t.y > yBot + 3)

  // Column widths in twips from the PDF's real point widths; only scaled DOWN
  // if the table is wider than the page text column (keeps real proportions and
  // avoids stretching a narrow table across the page).
  let colWidths = colXs.slice(0, -1).map((x, c) => (colXs[c + 1] - x) * TWIPS_PER_PT)
  const totalW = colWidths.reduce((s, w) => s + w, 0)
  if (totalW > TEXT_WIDTH_TWIPS) colWidths = colWidths.map((w) => (w / totalW) * TEXT_WIDTH_TWIPS)

  // Split any tall, unruled band into its logical rows using the text positions.
  const refinedRowYs = refineRowYs(rowYs, colXs, tableTexts)

  const tableRows = []
  for (let r = 0; r < refinedRowYs.length - 1; r++) {
    const row = []
    for (let c = 0; c < colXs.length - 1; c++) {
      const cx0 = colXs[c]
      const cx1 = colXs[c + 1]
      const ry0 = refinedRowYs[r]
      const ry1 = refinedRowYs[r + 1]
      const cellTexts = tableTexts.filter((t) => t.x + 2 >= cx0 && t.x < cx1 && t.y >= ry0 - 3 && t.y < ry1)
      const lines = linesOf(cellTexts).map((l) => [run(l.text, l.bold, l.size)])
      row.push({ lines })
    }
    tableRows.push(row)
  }

  return [
    ...paragraphsBlocks(above, g.width),
    { kind: 'table', colWidths, rows: tableRows },
    ...paragraphsBlocks(below, g.width)
  ]
}

/**
 * Turn the extracted per-page geometry into the doc-model blocks (paragraphs +
 * real tables) that core/docxBuilder writes as an editable, Word-valid .docx.
 * A hard page break is inserted before every page after the first.
 */
export function geometryToBlocks(pages: PageGeometry[]): DocBlock[] {
  const out: DocBlock[] = []
  pages.forEach((page, i) => {
    const blocks = pageBlocks(page)
    if (i > 0) {
      // Carry the page break on the first block if it's a paragraph, else add one.
      if (blocks[0]?.kind === 'paragraph') blocks[0].pageBreak = true
      else out.push({ kind: 'paragraph', runs: [], pageBreak: true })
    }
    out.push(...blocks)
  })
  return out.length ? out : [{ kind: 'paragraph', runs: [] }]
}

/** Build simple paragraph blocks from plain OCR text lines (the photos → Word path, which has no layout to reconstruct). */
export function textLinesToBlocks(lines: string[]): DocBlock[] {
  const blocks: DocBlock[] = lines.map((line) => ({ kind: 'paragraph', runs: line.trim() ? [{ text: line }] : [] }))
  return blocks.length ? blocks : [{ kind: 'paragraph', runs: [] }]
}

/** Whether the geometry has any extractable text — false for a scanned/image PDF (which must fall back to OCR or the LibreOffice image conversion). */
export function geometryHasText(pages: PageGeometry[]): boolean {
  return pages.some((p) => p.texts.length > 0)
}

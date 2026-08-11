import type { DocBlock, DocRun } from './docxBuilder'

/**
 * Offline "AI" table reconstruction from the OCR detector's own output: every
 * recognised text line comes with a bounding box (see electron/ocr.ts's
 * OcrLine.box), so a page image — a photo, a scan, or a rendered PDF page — can
 * be rebuilt into rows and columns purely from where the text sits, with no
 * extra model and nothing leaving the machine. Works where the PDF text-layer
 * approach can't: photos, scans, and pages whose embedded fonts are broken
 * (e.g. Telugu rendered as glyph fragments), because it reads the image.
 */

export interface OcrLineBox {
  text: string
  /** Line box in image pixels. */
  x: number
  y: number
  w: number
  h: number
}
export interface OcrPage {
  lines: OcrLineBox[]
}

const TEXT_WIDTH_TWIPS = 9360

const median = (nums: number[]): number => {
  if (nums.length === 0) return 0
  const s = [...nums].sort((a, b) => a - b)
  return s[Math.floor(s.length / 2)]
}

/** Collapse near-equal values (within tol) into single averaged, sorted values. */
function cluster(vals: number[], tol: number): number[] {
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
  return groups.map((g) => g.c)
}

/** A run whose font size is scaled from the line's box height relative to the page's median line height. */
function run(text: string, medH: number, h: number): DocRun {
  const size = Math.max(12, Math.min(56, Math.round(22 * (h / (medH || h || 12)))))
  return { text, size }
}

interface Row {
  yc: number
  lines: OcrLineBox[]
}

/** Group lines into visual rows by their vertical centre. */
function groupRows(lines: OcrLineBox[], tol: number): Row[] {
  const sorted = [...lines].sort((a, b) => a.y + a.h / 2 - (b.y + b.h / 2))
  const rows: Row[] = []
  for (const l of sorted) {
    const yc = l.y + l.h / 2
    const last = rows[rows.length - 1]
    if (last && Math.abs(yc - last.yc) <= tol) {
      last.lines.push(l)
      last.yc = (last.yc * (last.lines.length - 1) + yc) / last.lines.length
    } else {
      rows.push({ yc, lines: [l] })
    }
  }
  for (const r of rows) r.lines.sort((a, b) => a.x - b.x)
  return rows
}

/** Which column (index into colStarts) a line's left edge falls in. */
function colOf(x: number, colStarts: number[]): number {
  let c = 0
  for (let i = 0; i < colStarts.length; i++) if (x + 3 >= colStarts[i]) c = i
  return c
}

function pageBlocks(page: OcrPage): DocBlock[] {
  const lines = page.lines.filter((l) => l.text.trim())
  if (lines.length === 0) return []
  const left = Math.min(...lines.map((l) => l.x))
  const right = Math.max(...lines.map((l) => l.x + l.w))
  const medH = median(lines.map((l) => l.h)) || 12
  const rows = groupRows(lines, medH * 0.7)

  // Columns from the clustered left edges of every line. A form's label lines
  // and value lines each share a left edge, so this recovers the columns.
  const colStarts = cluster(
    lines.map((l) => l.x),
    Math.max(medH * 1.2, (right - left) * 0.02)
  )

  // Not tabular (one column, or almost every row is a single line): paragraphs.
  const multiColRows = rows.filter((r) => new Set(r.lines.map((l) => colOf(l.x, colStarts))).size >= 2).length
  if (colStarts.length < 2 || multiColRows < 2) {
    return rows.map((r) => ({
      kind: 'paragraph' as const,
      runs: [run(r.lines.map((l) => l.text).join(' '), medH, medH)]
    }))
  }

  const bounds = [...colStarts, right + 1]
  const rawWidths = bounds.slice(0, -1).map((b, i) => bounds[i + 1] - b)
  const totalW = rawWidths.reduce((s, w) => s + w, 0) || 1
  const colWidths = rawWidths.map((w) => (w / totalW) * TEXT_WIDTH_TWIPS)

  const tableRows = rows.map((r) =>
    colStarts.map((_, ci) => {
      const cellLines = r.lines.filter((l) => colOf(l.x, colStarts) === ci)
      if (cellLines.length === 0) return { lines: [] as DocRun[][] }
      const text = cellLines.map((l) => l.text).join(' ').replace(/\s+/g, ' ').trim()
      const h = median(cellLines.map((l) => l.h))
      return { lines: text ? [[run(text, medH, h)]] : [] }
    })
  )

  return [{ kind: 'table', colWidths, rows: tableRows }]
}

/** Reconstruct the OCR'd pages into doc-model blocks (tables/paragraphs) for the Word builder, page break between pages. */
export function layoutToBlocks(pages: OcrPage[]): DocBlock[] {
  const out: DocBlock[] = []
  pages.forEach((page, i) => {
    const blocks = pageBlocks(page)
    if (i > 0) {
      if (blocks[0]?.kind === 'paragraph') blocks[0].pageBreak = true
      else out.push({ kind: 'paragraph', runs: [], pageBreak: true })
    }
    out.push(...blocks)
  })
  return out.length ? out : [{ kind: 'paragraph', runs: [] }]
}

/** Reconstruct the OCR'd pages into a flat grid of rows for the Excel export (one output row per visual row, cells by column). */
export function layoutToRows(pages: OcrPage[]): string[][] {
  const out: string[][] = []
  for (const page of pages) {
    const lines = page.lines.filter((l) => l.text.trim())
    if (lines.length === 0) continue
    const right = Math.max(...lines.map((l) => l.x + l.w))
    const left = Math.min(...lines.map((l) => l.x))
    const medH = median(lines.map((l) => l.h)) || 12
    const rows = groupRows(lines, medH * 0.7)
    const colStarts = cluster(
      lines.map((l) => l.x),
      Math.max(medH * 1.2, (right - left) * 0.02)
    )
    for (const r of rows) {
      const cells = colStarts.map((_, ci) =>
        r.lines
          .filter((l) => colOf(l.x, colStarts) === ci)
          .map((l) => l.text)
          .join(' ')
          .trim()
      )
      out.push(cells)
    }
  }
  return out
}

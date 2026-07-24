export interface OcrBox {
  x0: number
  y0: number
  x1: number
  y1: number
}

export interface OcrWord {
  text: string
  bbox: OcrBox
}

/**
 * The dividing line between "ordinary spacing between two words on the
 * same line" and "a real gap between two table columns", found from the
 * gaps themselves rather than assumed as a fixed multiple of some average
 * (which a single unusually wide word/description can throw off). Sorts
 * every observed gap and finds the biggest *relative* jump between
 * consecutive gap sizes — that jump is the natural break between the small,
 * fairly uniform in-line spacing and the much larger column gaps.
 *
 * Requires that jump to be a meaningful one (>1.4x, not just technically
 * bigger) before trusting it — a row of short, evenly-spaced single-word
 * cells (e.g. "1  25.50  350.00") has no small "same word" gaps to find at
 * all, and treating its one modest gap-size difference as the split point
 * would wrongly merge separate columns together. Without a real jump,
 * nothing is treated as "same column" — safer to leave a table over-split
 * into extra cells (easy to notice and merge by hand during review) than
 * to silently combine two different columns' values into one.
 */
function findColumnGapThreshold(positiveGaps: number[]): number {
  if (positiveGaps.length < 2) return 0

  const sorted = [...positiveGaps].sort((a, b) => a - b)
  const MEANINGFUL_JUMP = 1.4
  let bestRatio = MEANINGFUL_JUMP
  let bestIdx = -1
  for (let i = 1; i < sorted.length; i++) {
    const ratio = sorted[i] / Math.max(sorted[i - 1], 1)
    if (ratio > bestRatio) {
      bestRatio = ratio
      bestIdx = i
    }
  }
  if (bestIdx === -1) return 0
  return (sorted[bestIdx - 1] + sorted[bestIdx]) / 2
}

/**
 * Rebuild a grid of table cells (rows x columns) from OCR word-level
 * positions — Tesseract only knows individual words and where they sit on
 * the page, not which "cell" of a table each belongs to. This is a
 * best-effort reconstruction, not a guaranteed-correct one: a skewed photo,
 * uneven lighting, or an unusually laid-out table can still throw off the
 * row/column grouping — that's expected and why the extracted result is
 * shown for review before export, not saved directly.
 *
 * Rows: words are grouped by vertical (Y-center) proximity — words whose
 * Y-centers are within `rowTolerance` (a fraction of the average word
 * height) of each other are treated as being on the same printed line.
 *
 * Columns: words are sorted left to right and merged into column "bands"
 * by the actual whitespace between one word's right edge and the next
 * word's left edge (not the distance between their left edges, which would
 * misjudge a wide word followed immediately by a narrow one as a big gap).
 * See findColumnGapThreshold for how the same-column vs. new-column cutoff
 * is chosen. Several words landing in the same row+column (e.g. a wrapped
 * multi-word description) are joined with a space.
 */
export function reconstructGrid(words: OcrWord[], rowTolerance = 0.6): string[][] {
  if (words.length === 0) return []

  const avgHeight = words.reduce((sum, w) => sum + (w.bbox.y1 - w.bbox.y0), 0) / words.length
  const yTolerance = avgHeight * rowTolerance

  interface Row {
    words: OcrWord[]
    yCenter: number
  }
  const rows: Row[] = []
  const sortedByY = [...words].sort((a, b) => a.bbox.y0 - b.bbox.y0)
  for (const w of sortedByY) {
    const yc = (w.bbox.y0 + w.bbox.y1) / 2
    const row = rows.find((r) => Math.abs(r.yCenter - yc) <= yTolerance)
    if (row) {
      row.yCenter = (row.yCenter * row.words.length + yc) / (row.words.length + 1)
      row.words.push(w)
    } else {
      rows.push({ words: [w], yCenter: yc })
    }
  }
  rows.sort((a, b) => a.yCenter - b.yCenter)
  for (const r of rows) r.words.sort((a, b) => a.bbox.x0 - b.bbox.x0)

  const sortedByX = [...words].sort((a, b) => a.bbox.x0 - b.bbox.x0)
  const positiveGaps: number[] = []
  for (let i = 1; i < sortedByX.length; i++) {
    const gap = sortedByX[i].bbox.x0 - sortedByX[i - 1].bbox.x1
    if (gap > 0) positiveGaps.push(gap)
  }
  const columnThreshold = findColumnGapThreshold(positiveGaps)

  interface Band {
    start: number
    end: number
  }
  const bands: Band[] = []
  const bandIndexByWord = new Map<OcrWord, number>()
  for (const w of sortedByX) {
    const last = bands[bands.length - 1]
    if (last && w.bbox.x0 - last.end <= columnThreshold) {
      last.end = Math.max(last.end, w.bbox.x1)
      bandIndexByWord.set(w, bands.length - 1)
    } else {
      bands.push({ start: w.bbox.x0, end: w.bbox.x1 })
      bandIndexByWord.set(w, bands.length - 1)
    }
  }

  return rows.map((row) => {
    const cells: string[] = []
    for (const w of row.words) {
      const col = bandIndexByWord.get(w) ?? 0
      cells[col] = cells[col] ? `${cells[col]} ${w.text}` : w.text
    }
    return Array.from({ length: cells.length }, (_, i) => cells[i] ?? '')
  })
}

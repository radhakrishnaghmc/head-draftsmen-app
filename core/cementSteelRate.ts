/**
 * Reads the Engineer-in-Chief (PH)'s periodic "Cement & Steel Rates" circular
 * — a scanned PDF (no text layer), OCR'd the same way as an estimate photo
 * (see src/pdfToImages.ts's pdfPagesToDataUrls + api.ocrEstimatePhotos) — to
 * fill the TS Note's cement/steel/M.S.Flats rate sentence. The OCR engine
 * returns the whole rate table one PRINTED ROW per line (e.g. "a) Ordinary
 * Portland Cement(43/53 grade) 5,100/- M.T"), so each field is read off a
 * single matched line rather than needing column reconstruction.
 */

import { CEMENT_STEEL_MONTHS } from './cementSteelRates'

export interface CementSteelRates {
  /** "146/Cement & Steel/T3/May/2025-26" — the bit after "C. Memo.No." */
  memoRef?: string
  /** "29.05.2026" — the circular's own Dated: line, normalised to dd.mm.yyyy. */
  memoDate?: string
  /** "May-2026" — the rate period the circular covers. */
  monthYear?: string
  /** Ordinary Portland Cement (43/53 grade) rate, e.g. "5,100". */
  cementRate?: string
  /** Category-B (TMT/HYSD) Reinforcement Steel rate, e.g. "50,000" — matches the office's own sample notes, which always cite the Category-B figure. */
  steelRate?: string
  /** M.S. Flats rate, e.g. "55,000". */
  msFlatsRate?: string
}

/** The trailing "5,100/- M.T" (or "5,100/-") rate off one OCR'd table row. */
function rateOnLine(line: string | undefined): string | undefined {
  if (!line) return undefined
  const m = /([\d][\d,]*)\s*\/-/.exec(line)
  return m ? m[1] : undefined
}

/**
 * Find the first line matching `labelRe` and read the rate off it — or, if
 * the OCR happened to split that printed row into two detected lines (a
 * lower-DPI render can draw the box boundaries differently than a same-page
 * higher-DPI test render), off one of the next couple of lines instead. Every
 * material's rate lookup uses this so none of them depend on OCR merging a
 * label and its rate onto exactly one line.
 */
function rateNear(lines: string[], labelRe: RegExp, span = 3): string | undefined {
  const idx = lines.findIndex((l) => labelRe.test(l))
  if (idx === -1) return undefined
  for (let i = idx; i < Math.min(idx + span, lines.length); i++) {
    const r = rateOnLine(lines[i])
    if (r) return r
  }
  return undefined
}

/** dd-mm-yyyy (the circular's own date style) -> dd.mm.yyyy. */
function normaliseDate(d: string, m: string, y: string): string {
  return `${d.padStart(2, '0')}.${m.padStart(2, '0')}.${y}`
}

/**
 * Parse the OCR'd line list (already in top-to-bottom reading order, one
 * table row per line) into the rate fields the TS Note needs. Every field is
 * optional — a line the OCR missed or mis-read just leaves that field blank
 * for the office to fill by hand, same convention as the app's other
 * hand-fillable references.
 */
export function parseCementSteelRateLines(lines: string[]): CementSteelRates {
  const text = lines.join('\n')
  const out: CementSteelRates = {}

  const memoMatch = /Memo\.?\s*No\.?\s*([\d]+\/[^\n]*?\/\d{4}-\d{2})[^\n]*?Dated:?\s*(\d{1,2})[-./](\d{1,2})[-./](\d{4})/i.exec(
    text
  )
  if (memoMatch) {
    out.memoRef = memoMatch[1].trim()
    out.memoDate = normaliseDate(memoMatch[2], memoMatch[3], memoMatch[4])
  }

  const monthMatch = /for the month of\s+([A-Za-z]+-\d{4})/i.exec(text)
  if (monthMatch) out.monthYear = monthMatch[1]

  // "Ordinary Portland Cement" is the primary anchor; if OCR garbled that
  // specific label, fall back to the "1 CEMENT" section header — either way,
  // the rate itself is read from whichever of the next few lines has one.
  out.cementRate =
    rateNear(lines, /Ordinary\s+Portland\s+Cement/i) ?? rateNear(lines, /^1\b.{0,3}CEMENT\b/i)

  out.msFlatsRate = rateNear(lines, /M\.?\s*S\.?\s*Flats/i)

  // Category-B's rate sits on the row right after the "Category-B Steel"
  // sub-heading (that heading line itself carries no rate).
  out.steelRate = rateNear(lines, /Category-?B\s+Steel/i)

  return out
}

/** A "Cement (July 2026)= Rs.5200/MT" / "Steel HYSD(July 2026) =Rs. 52800/MT" style cell matched in an uploaded estimate's raw grid (see TS FOR KOMPALLY.xlsx's header rows) — one such line per material, giving the month/year whose departmental circular Replicate Technical Sanction should auto-fetch. */
export interface CementSteelRateCellRef {
  row: number
  col: number
  sheetName: string
  kind: 'cement' | 'steel'
  /** Lowercase month name, one of CEMENT_STEEL_MONTHS. */
  month: string
  /** The year the estimate itself cites, or the current year when the estimate names only a month (never guessed from anywhere else). */
  year: number
  /** The exact old-rate substring matched (e.g. "5200") — replaced in place with the newly fetched rate so the rest of the cell's wording/unit is left untouched. */
  oldRateText: string
  /** Full original cell text. */
  raw: string
}

// "Cement (July 2026)= Rs.5200/MT" / "Steel HYSD(July 2026) =Rs. 52800/MT" —
// label, then month (year optional — estimates sometimes name only the
// month), then the "Rs." rate figure, wherever it lands relative to the
// parenthesised month/year.
const CEMENT_STEEL_CELL_RE = /^(Cement|Steel(?:\s*HYSD)?)\b.*?\(?\s*([A-Za-z]+)\.?\s*(\d{4})?\)?.*?Rs\.?\s*([\d][\d,]*(?:\.\d+)?)/i

/**
 * Scans an uploaded estimate's raw cells for the office's own
 * "Cement (<Month> <Year>)= Rs.<rate>/MT" / "Steel HYSD(<Month> <Year>)
 * =Rs.<rate>/MT" header convention, so Replicate Technical Sanction can
 * auto-fetch that month's departmental circular and refresh the figure
 * instead of silently carrying over whatever rate the old sanctioned
 * estimate happened to cite.
 */
export function findCementSteelRateCells(grids: { sheetName: string; grid: string[][] }[]): CementSteelRateCellRef[] {
  const out: CementSteelRateCellRef[] = []
  const currentYear = new Date().getFullYear()
  for (const g of grids) {
    for (let row = 0; row < g.grid.length; row++) {
      const gridRow = g.grid[row] ?? []
      for (let col = 0; col < gridRow.length; col++) {
        const raw = String(gridRow[col] ?? '').trim()
        if (!raw) continue
        const m = CEMENT_STEEL_CELL_RE.exec(raw)
        if (!m) continue
        const month = m[2].toLowerCase()
        if (!CEMENT_STEEL_MONTHS.includes(month)) continue
        out.push({
          row,
          col,
          sheetName: g.sheetName,
          kind: /^cement/i.test(m[1]) ? 'cement' : 'steel',
          month,
          year: m[3] ? Number(m[3]) : currentYear,
          oldRateText: m[4],
          raw
        })
      }
    }
  }
  return out
}

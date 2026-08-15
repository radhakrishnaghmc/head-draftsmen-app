/**
 * Reads the Engineer-in-Chief (PH)'s periodic "Cement & Steel Rates" circular
 * — a scanned PDF (no text layer), OCR'd the same way as an estimate photo
 * (see src/pdfToImages.ts's pdfPagesToDataUrls + api.ocrEstimatePhotos) — to
 * fill the TS Note's cement/steel/M.S.Flats rate sentence. The OCR engine
 * returns the whole rate table one PRINTED ROW per line (e.g. "a) Ordinary
 * Portland Cement(43/53 grade) 5,100/- M.T"), so each field is read off a
 * single matched line rather than needing column reconstruction.
 */

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

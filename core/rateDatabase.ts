export interface RateEntry {
  description: string
  rate: string
  sheet: string
  /** The raw material/labour/machinery breakdown rows that produced this rate — for reproducing as a rate-analysis backup sheet in the generated Technical Sanction. */
  breakdown: string[][]
}

function isNumericLike(s: string): boolean {
  const t = s.trim().replace(/,/g, '')
  return t !== '' && Number.isFinite(Number(t))
}

// Matches the label cell of a block's final costed line — "Grand Total",
// "Rate per cum = (a+b+c)/10", "Total (Rs)", "Basic rate per Cum", etc. The
// wording varies across sheets, but every variant contains "total" or "rate".
const FINAL_LABEL_RE = /\btotal\b|\brate\b/i

// Some sheets present a handful of concrete-grade variants (M10/M15/M20/
// M25/M30/M40) as a small side table instead of separate "Rate per…" lines —
// the grade token stands in for a label, with the applicable rate as the
// row's last populated number (the columns before it are the intermediate
// cement-quantity/cost workings that produced it).
const GRADE_RE = /^m\d{1,3}$/i

// The rate value sits somewhere to the right of the label on the same row,
// usually the very next populated cell — take the first one, not the last,
// since a stray reference number can sit far to the right (e.g. a Remarks
// column) after several blank cells.
function findRateAfter(cells: string[], labelIdx: number): string | null {
  for (let i = labelIdx + 1; i < cells.length; i++) {
    if (cells[i] !== '' && isNumericLike(cells[i])) return cells[i]
  }
  return null
}

/**
 * Extract {description, rate} pairs from one sheet of a rates database
 * ("Telangana/Andhra Pradesh Standard Data" style workbooks — Building Data,
 * RMC, the RBR-* chapters, etc.). Each item spans a block of rows: a lead row
 * carrying an index code and/or serial number together with the item's
 * description (position varies sheet to sheet, so it's found by picking the
 * longest text cell in that row rather than a fixed column), a material/
 * labour/machinery cost breakdown, and a line further down labelled "Grand
 * Total" / "Rate per <unit> = …" / etc. carrying the final worked-out rate.
 *
 * A new block starts wherever column 0 holds an index code, or column 1
 * holds a bare serial number — both are blank on every other row in the
 * sheet. Sheets that follow neither convention (a handful of the workbook's
 * 40+ tabs use an entirely ad-hoc layout) simply yield nothing here; that's
 * an acceptable gap rather than a wrong guess.
 */
export function extractRateEntriesFromGrid(grid: string[][], sheetName: string): RateEntry[] {
  const entries: RateEntry[] = []
  let pendingDescription: string | null = null
  // Rows seen since the start of the current item (or since the last
  // variant's rate line, for a multi-variant block) — becomes that variant's
  // `breakdown`. Reset after every match so each variant only carries its
  // own specific buildup, not the accumulated history of prior variants.
  let blockRows: string[][] = []

  for (const raw of grid) {
    const cells = raw.map((c) => String(c ?? '').trim())
    const nonEmpty = cells.filter((c) => c !== '')
    if (nonEmpty.length === 0) continue

    const isItemStart = cells[0] !== '' || /^\d+$/.test(cells[1] ?? '')
    if (isItemStart) {
      const longest = nonEmpty.reduce((a, b) => (b.length > a.length ? b : a), '')
      pendingDescription = longest.length >= 8 ? longest : null
      blockRows = pendingDescription ? [cells] : []
      continue
    }

    // Some sheets give the item-start row only a short title (e.g.
    // "Excavation for Structures") and put the real, much longer
    // description on its own line right after — a lone text cell with
    // nothing else on the row. Prefer whichever candidate is longer, and —
    // when nothing is pending yet — let it seed pendingDescription outright:
    // a handful of sheets never give their top description an index code or
    // serial number at all (it just sits alone on its own line), and a
    // spurious numeric-looking row above it (e.g. a "1  2  3  4  5" column-
    // index line) can otherwise null out anything pending before this point.
    if (nonEmpty.length === 1 && /[a-zA-Z]/.test(nonEmpty[0]) && nonEmpty[0].length >= 8) {
      if (!pendingDescription || nonEmpty[0].length > pendingDescription.length) {
        pendingDescription = nonEmpty[0]
      }
    }

    if (!pendingDescription) continue
    blockRows.push(cells)

    const labelIdx = cells.findIndex((c) => FINAL_LABEL_RE.test(c))
    if (labelIdx !== -1) {
      const rate = findRateAfter(cells, labelIdx)
      // Deliberately don't clear pendingDescription after a match: many
      // items have several total/rate lines under one description (e.g.
      // manual vs. mechanical means × depth range, all sharing the same
      // paragraph) — each one is a genuine, distinct rate for that
      // description, only distinguished by structural sub-labels this
      // parser doesn't track. Keeping the description alive until the next
      // item-start row lets every variant surface as its own entry, which is
      // what makes the multi-rate/ambiguous case detectable downstream
      // instead of silently keeping just the first.
      if (rate) entries.push({ description: pendingDescription, rate, sheet: sheetName, breakdown: blockRows })
      blockRows = []
      continue
    }

    const gradeIdx = cells.findIndex((c) => GRADE_RE.test(c))
    if (gradeIdx !== -1) {
      const numeric = cells.filter((c) => c !== '' && isNumericLike(c))
      const rate = numeric[numeric.length - 1]
      if (rate) {
        entries.push({
          description: `${pendingDescription} ${cells[gradeIdx].toUpperCase()}`,
          rate,
          sheet: sheetName,
          breakdown: blockRows
        })
      }
      blockRows = []
    }
  }

  return entries
}

/** Flatten every sheet of a rates database workbook into one searchable list. */
export function buildRateIndex(sheets: { name: string; grid: string[][] }[]): RateEntry[] {
  return sheets.flatMap((s) => extractRateEntriesFromGrid(s.grid, s.name))
}

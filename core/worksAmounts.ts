// General, app-wide money rules for the Works List's estimate/EMD/ASD/
// Contract Amount columns — applied uniformly wherever a generated document
// fills one of these placeholders (the Bid Document generator, and any
// user-authored {{Placeholder}} template filled from a Works List row).
import type { ExcelTable } from './types'
import { rankByEmbedding } from './embeddingMatch'

// A match below this score is treated as "no real match" — same threshold
// used for placeholder/column matching elsewhere (core/columnMatch.ts).
const EMBEDDING_THRESHOLD = 0.5

/** "1" or ".30" (Lakhs) -> 100000 or 30000 (rupees). "Amount of estimate" is entered on the Works List in Lakhs. */
export function lakhsToRupees(lakhs: string): number {
  const n = Number(String(lakhs).replace(/,/g, '').trim())
  return Number.isFinite(n) ? Math.round(n * 100000) : 0
}

/**
 * The tender percentage cell from a Works List row. It's the standard "Tender
 * Percentage" column, but the same figure is sometimes headed "TP" or "Quoted
 * Commission" (e.g. on an imported sheet), so those aliases are read as
 * fallbacks — the standard column wins when both are present and filled.
 */
export function tenderPercentFromRow(row: Record<string, string>): string {
  const std = (row['Tender Percentage'] ?? '').trim()
  if (std) return std
  for (const key of Object.keys(row)) {
    const k = key.trim().toLowerCase()
    if (k === 'tp' || k === 'quoted commission') {
      const v = (row[key] ?? '').trim()
      if (v) return v
    }
  }
  return ''
}

/**
 * Parse a rupee-denominated Works List cell to a whole number of rupees.
 * ECV and Contract Amount are stored on the Works List in rupees (the tender
 * portal reports them that way), unlike "Amount of estimate" which is in
 * Lakhs — so those columns are read with this, not lakhsToRupees. Blank or
 * unparseable -> 0.
 */
export function rupeesFromCell(cell: string): number {
  const n = Number(String(cell).replace(/,/g, '').trim())
  return Number.isFinite(n) ? Math.round(n) : 0
}

/** A rupee amount as a plain whole-number string for storing in a rupee-denominated Works List cell (ECV / Contract Amount). */
export function rupeesToCell(rupees: number): string {
  return String(Math.round(rupees))
}

/** 1234567 -> "12,34,567" (Indian digit grouping: last 3 digits, then pairs). */
export function indianDigitGroups(n: number): string {
  const sign = n < 0 ? '-' : ''
  const s = String(Math.round(Math.abs(n)))
  const lastThree = s.slice(-3)
  const other = s.slice(0, -3)
  const grouped = other === '' ? '' : other.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ','
  return `${sign}${grouped}${lastThree}`
}

/** "Rs 1,00,000/-" */
export function formatRupees(n: number): string {
  return `Rs ${indianDigitGroups(n)}/-`
}

/** "18%", "18", " -5 " -> 18, 18, -5. Blank/unparseable -> undefined ("not available"), not 0. */
function parsePercent(v: string | undefined): number | undefined {
  const s = String(v ?? '').replace(/[%,\s]/g, '')
  if (s === '') return undefined
  const n = Number(s)
  return Number.isFinite(n) ? n : undefined
}

export interface ComputedAmounts {
  /** "Amount of estimate" / "Estimate Amount", converted from Lakhs to rupees. */
  estimate: number
  /** "ECV" in rupees (stored in rupees, so read as-is) — null when ECV is blank. Never falls back to `estimate`: the two are distinct figures and must not be conflated. */
  ecv: number | null
  /** 1% of ECV — null when ECV isn't available yet, rather than computed off the estimate. */
  emd1: number | null
  /** 1.5% of ECV — null when ECV isn't available yet, rather than computed off the estimate. */
  emd1_5: number | null
  /** (Tender Percentage - 25%) of ECV, only once Tender Percentage exceeds 25% — null when ECV isn't available yet, otherwise 0 below the 25% threshold. */
  asd: number | null
  /** ECV net of the tendered percentage: ECV * (1 - Tender Percentage) — null when Tender Percentage or ECV isn't available yet, rather than assuming 0%. */
  contractAmount: number | null
}

/**
 * Derive every amount-related figure from a Works List row's raw columns:
 * - Estimate/ECV: Lakhs -> rupees.
 * - EMD @ 1%/1.5%: percentages of ECV.
 * - ASD (Additional Security Deposit): charged only once Tender Percentage
 *   exceeds 25%, at (Tender Percentage - 25%) of ECV.
 * - Contract Amount: ECV net of the tendered percentage — left null (not
 *   0%-tendered) when Tender Percentage isn't on the row at all, since a
 *   contract amount doesn't exist before a tender percentage is quoted.
 *
 * ECV and "Amount of estimate" are distinct figures (ECV only exists once a
 * tender is held) — every ECV-derived figure below is left null, never
 * computed off the estimate instead, when a row's ECV cell is still blank.
 */
export function computeWorkAmounts(row: Record<string, string>): ComputedAmounts {
  const estimate = lakhsToRupees(row['Amount of estimate'] ?? '')
  const ecvRaw = row['ECV']
  // ECV is stored in rupees (unlike Amount of estimate, in Lakhs).
  const ecv = ecvRaw?.trim() ? rupeesFromCell(ecvRaw) : null
  const tenderPercent = parsePercent(tenderPercentFromRow(row))

  const emd1 = ecv !== null ? Math.round(ecv * 0.01) : null
  const emd1_5 = ecv !== null ? Math.round(ecv * 0.015) : null
  const asd =
    ecv === null
      ? null
      : tenderPercent !== undefined && tenderPercent > 25
        ? Math.round(ecv * ((tenderPercent - 25) / 100))
        : 0
  const contractAmount =
    ecv !== null && tenderPercent !== undefined ? Math.round(ecv * (1 - tenderPercent / 100)) : null

  return { estimate, ecv, emd1, emd1_5, asd, contractAmount }
}

/**
 * An enriched copy of a Works List row where every amount-bearing column
 * (and the "Estimate Amount" wording some templates use in place of "Amount
 * of estimate") resolves to its computed, Indian-formatted value
 * ("Rs 1,00,000/-") instead of the raw spreadsheet figure. Contract Amount is
 * left blank when Tender Percentage isn't available.
 */
export function withComputedAmounts(row: Record<string, string>): Record<string, string> {
  const c = computeWorkAmounts(row)
  // A blank "Amount of estimate" cell stays blank — never rendered as "Rs 0/-"
  // (e.g. an Issue Document previewed/issued with no work selected). Only a
  // filled cell gets the formatted "Rs …/-" figure.
  const estimate = (row['Amount of estimate'] ?? '').trim() ? formatRupees(c.estimate) : ''
  return {
    ...row,
    'Amount of estimate': estimate,
    'Estimate Amount': estimate,
    ECV: c.ecv !== null ? formatRupees(c.ecv) : '',
    'EMD 1%': c.emd1 !== null ? formatRupees(c.emd1) : '',
    'EMD 1.5%': c.emd1_5 !== null ? formatRupees(c.emd1_5) : '',
    ASD: c.asd !== null ? formatRupees(c.asd) : '',
    'Contract Amount': c.contractAmount !== null ? formatRupees(c.contractAmount) : ''
  }
}

/**
 * 2500000 (rupees) -> "25"; 320000 -> "3.2"; 37500 -> "0.375" — a plain
 * Lakhs figure matching how the Works List's own amount columns are
 * entered. Rounds to 5 decimal places purely to kill floating-point noise
 * (rupees are always whole numbers, so their Lakhs value never needs more
 * than 5) — it does not lose precision on the rupee amount.
 */
export function rupeesToLakhsString(rupees: number): string {
  return String(Math.round((rupees / 100000) * 100000) / 100000)
}

export interface EcvMatchResult {
  table: ExcelTable
  /** Whether a Works List row's "Name of the work" matched (exactly, or via embeddings — see matchedViaAi). */
  matched: boolean
  /** True when the match came from semantic similarity (embeddings) rather than an exact name match — worth a "please verify" flag, same as column-matching elsewhere. */
  matchedViaAi?: boolean
}

/**
 * Find the Works List row whose "Name of the work" matches (case- and
 * whitespace-insensitive) `workName`, and return an updated table with that
 * row's ECV, EMD 1% and EMD 1.5% filled in from `ecvRupees` (a BOQ/estimate
 * item total) — in rupees, matching how ECV is stored (EMD written in rupees
 * too so it stays consistent with the rupee ECV it's derived from). Returns
 * the table unchanged (`matched: false`) if
 * there's no "Name of the work" column, no name to match, or no row matches —
 * callers should treat that as "nothing to update", not an error.
 *
 * An estimate's title-block work name very often differs slightly in wording
 * from the Works List's own entry (abbreviations, punctuation, rephrasing) —
 * when the exact match fails and `embeddings` are supplied (the same
 * work-name text and every candidate row name, both embedded by the caller),
 * the closest semantic match above the threshold is used instead.
 */
export function applyEcvFromBoq(
  table: ExcelTable,
  workName: string,
  ecvRupees: number,
  embeddings?: { workNameVector: number[]; rowNameVectors: number[][] }
): EcvMatchResult {
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')
  const target = norm(workName)
  if (!nameHeader || !target) return { table, matched: false }

  let idx = table.rows.findIndex((r) => norm(r[nameHeader] ?? '') === target)
  let matchedViaAi = false
  if (idx === -1 && embeddings) {
    const [best] = rankByEmbedding(embeddings.workNameVector, embeddings.rowNameVectors)
    if (best && best.score >= EMBEDDING_THRESHOLD) {
      idx = best.index
      matchedViaAi = true
    }
  }
  if (idx === -1) return { table, matched: false }

  const emd1 = Math.round(ecvRupees * 0.01)
  const emd1_5 = Math.round(ecvRupees * 0.015)
  const rows = table.rows.map((r, i) =>
    i !== idx
      ? r
      : {
          ...r,
          ECV: rupeesToCell(ecvRupees),
          'EMD 1%': rupeesToCell(emd1),
          'EMD 1.5%': rupeesToCell(emd1_5)
        }
  )
  return { table: { ...table, rows }, matched: true, matchedViaAi }
}

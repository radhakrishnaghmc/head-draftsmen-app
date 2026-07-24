// General, app-wide money rules for the Works List's estimate/EMD/ASD/
// Contract Amount columns — applied uniformly wherever a generated document
// fills one of these placeholders (the Bid Document generator, and any
// user-authored {{Placeholder}} template filled from a Works List row).
import type { ExcelTable } from './types'

/** "1" or ".30" (Lakhs) -> 100000 or 30000 (rupees). Estimate/ECV figures on the Works List are always entered in Lakhs. */
export function lakhsToRupees(lakhs: string): number {
  const n = Number(String(lakhs).replace(/,/g, '').trim())
  return Number.isFinite(n) ? Math.round(n * 100000) : 0
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
  /** "Estimate Amount ECV", converted from Lakhs to rupees — falls back to `estimate` when ECV is blank. */
  ecv: number
  /** 1% of ECV. */
  emd1: number
  /** 1.5% of ECV. */
  emd1_5: number
  /** (Tender Percentage - 25%) of ECV, only once Tender Percentage exceeds 25% — otherwise 0. */
  asd: number
  /** ECV net of the tendered percentage: ECV * (1 - Tender Percentage) — null when Tender Percentage isn't available yet, rather than assuming 0%. */
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
 */
export function computeWorkAmounts(row: Record<string, string>): ComputedAmounts {
  const estimate = lakhsToRupees(row['Amount of estimate'] ?? '')
  const ecvRaw = row['Estimate Amount ECV']
  const ecv = ecvRaw?.trim() ? lakhsToRupees(ecvRaw) : estimate
  const tenderPercent = parsePercent(row['Tender Percentage'])

  const emd1 = Math.round(ecv * 0.01)
  const emd1_5 = Math.round(ecv * 0.015)
  const asd = tenderPercent !== undefined && tenderPercent > 25 ? Math.round(ecv * ((tenderPercent - 25) / 100)) : 0
  const contractAmount = tenderPercent !== undefined ? Math.round(ecv * (1 - tenderPercent / 100)) : null

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
  return {
    ...row,
    'Amount of estimate': formatRupees(c.estimate),
    'Estimate Amount': formatRupees(c.estimate),
    'Estimate Amount ECV': formatRupees(c.ecv),
    'EMD 1%': formatRupees(c.emd1),
    'EMD 1.5%': formatRupees(c.emd1_5),
    ASD: formatRupees(c.asd),
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
  /** Whether a Works List row's "Name of the work" matched (case/whitespace-insensitive). */
  matched: boolean
}

/**
 * Find the Works List row whose "Name of the work" matches (case- and
 * whitespace-insensitive) `workName`, and return an updated table with that
 * row's Estimate Amount ECV, EMD 1% and EMD 1.5% filled in from `ecvRupees`
 * (a BOQ/estimate item total) — Lakhs-formatted, matching every other amount
 * column's own convention. Returns the table unchanged (`matched: false`) if
 * there's no "Name of the work" column, no name to match, or no row matches —
 * callers should treat that as "nothing to update", not an error.
 */
export function applyEcvFromBoq(table: ExcelTable, workName: string, ecvRupees: number): EcvMatchResult {
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')
  const target = norm(workName)
  if (!nameHeader || !target) return { table, matched: false }

  const idx = table.rows.findIndex((r) => norm(r[nameHeader] ?? '') === target)
  if (idx === -1) return { table, matched: false }

  const emd1 = Math.round(ecvRupees * 0.01)
  const emd1_5 = Math.round(ecvRupees * 0.015)
  const rows = table.rows.map((r, i) =>
    i !== idx
      ? r
      : {
          ...r,
          'Estimate Amount ECV': rupeesToLakhsString(ecvRupees),
          'EMD 1%': rupeesToLakhsString(emd1),
          'EMD 1.5%': rupeesToLakhsString(emd1_5)
        }
  )
  return { table: { ...table, rows }, matched: true }
}

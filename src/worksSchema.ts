import type { ExcelTable } from '@core/types'
import type { PlaceholderMatch } from '@core/createDocument'

/**
 * The standard "works" schema for the tender/agreement workflow. These columns
 * come from the APP template Excel and define the built-in Work database — the
 * user fills rows in-app (no Excel upload needed).
 */
export const WORKS_COLUMNS: string[] = [
  'Zone',
  'Circle',
  'Circle number',
  'Wincode',
  'Name of the work',
  'Amount of estimate',
  'ECV',
  'Contract Amount',
  'Sanction By',
  'Tender notice',
  'Financial Year',
  'Tender Notice No',
  'Tender notice Date',
  'Tender ID',
  'Tender Percentage',
  'Name of the Agency',
  'Address of the agency',
  'Phone number of the agency',
  'TCV',
  'Intimation Date',
  'EMD 1.5%',
  'EMD 1%',
  'ASD',
  'Reservation',
  'Completion Period',
  'Agmt Date',
  'Technical Sanc No',
  'TS date'
]

/** Build a fresh, empty works database (one blank row to type into). */
export function createWorksTable(): ExcelTable {
  const blankRow: Record<string, string> = {}
  for (const h of WORKS_COLUMNS) blankRow[h] = ''
  return {
    id: `works-${Date.now()}`,
    name: 'Works database',
    path: '',
    headers: [...WORKS_COLUMNS],
    rows: [blankRow]
  }
}

// "Estimate Amount ECV" was this column's name before it was shortened to
// "ECV" — carried across here so a table saved under the old schema doesn't
// lose its EMD-driving figure the next time it's normalized, rather than
// starting that column over blank.
const LEGACY_ECV_COLUMN = 'Estimate Amount ECV'

// "CNO" was this column's name before it was renamed to the clearer "Circle
// number" (they mean the same thing). Carried across the same way so a table
// saved under the old schema — or an imported sheet still using "CNO" — keeps
// its circle number instead of starting that column over blank.
const LEGACY_CNO_COLUMN = 'CNO'

/**
 * Force a table onto the standard works schema: guarantee every standard
 * column exists, while keeping any *extra* columns already on the table
 * (e.g. ones added by a previous import via applyWorksSchemaWithMapping) —
 * only ever adding to the schema, never silently dropping columns the user
 * already has. Guarantees at least one (blank) row.
 */
export function applyWorksSchema(table: ExcelTable): ExcelTable {
  const hasEcv = table.headers.includes('ECV')
  const hasCircleNumber = table.headers.includes('Circle number')
  const extraHeaders = table.headers.filter(
    (h) => !WORKS_COLUMNS.includes(h) && h !== LEGACY_ECV_COLUMN && h !== LEGACY_CNO_COLUMN
  )
  const headers = [...WORKS_COLUMNS, ...extraHeaders]
  const rows = (table.rows.length > 0 ? table.rows : [{}]).map((row) => {
    const next: Record<string, string> = {}
    for (const h of headers) {
      if (h === 'ECV' && !hasEcv) next[h] = row[LEGACY_ECV_COLUMN] ?? ''
      else if (h === 'Circle number' && !hasCircleNumber) next[h] = row[LEGACY_CNO_COLUMN] ?? ''
      else next[h] = row[h] ?? ''
    }
    return next
  })
  return { ...table, path: '', headers, rows }
}

/** Persisted-state schema version at which ECV/Contract Amount moved from Lakhs to rupees (see migrateEcvContractToRupees). */
export const ECV_RUPEES_STATE_VERSION = 2

// ECV and Contract Amount are stored in rupees (matching the tender portal);
// "Amount of estimate" stays in Lakhs. These two columns are the ones
// migrated from the old Lakhs storage.
const RUPEE_COLUMNS = ['ECV', 'Contract Amount']

/**
 * One-time migration for a Works List saved when ECV/Contract Amount were
 * stored in Lakhs: multiply those two columns' numeric cells by 100000 to
 * convert them to rupees. Idempotency is the caller's job — run this only
 * when the loaded state predates ECV_RUPEES_STATE_VERSION (see App.tsx), or
 * values would be inflated again on every load. Blank/unparseable cells and
 * every other column (including Amount of estimate) are left untouched.
 */
// A value already in rupees is at least this large; a value still in Lakhs is
// far smaller (a Lakhs figure of 100000 would be ₹1000 crore, absurd for these
// works). So a cell at or above this is already rupees (or was corrupted by an
// earlier over-migration) and must NOT be multiplied again — a defensive guard
// so a version-tracking slip (e.g. cloud state that dropped the version) can
// never silently inflate ECV/Contract Amount by ×100000 on every launch.
const ALREADY_RUPEES_MIN = 100000

export function migrateEcvContractToRupees(table: ExcelTable): ExcelTable {
  const rows = table.rows.map((row) => {
    const next = { ...row }
    for (const col of RUPEE_COLUMNS) {
      const raw = (next[col] ?? '').trim()
      if (!raw) continue
      const n = Number(raw.replace(/,/g, ''))
      if (Number.isFinite(n) && n > 0 && n < ALREADY_RUPEES_MIN) next[col] = String(Math.round(n * 100000))
    }
    return next
  })
  return { ...table, rows }
}

// ₹100 crore — no single work in this municipal workflow costs this much, so an
// ECV/Contract Amount at or above it was inflated by an earlier over-migration
// (the version-tracking slip fixed in electron/firebaseSync.ts, which re-ran the
// Lakhs->rupees ×100000 on every launch). Repairable because each extra run only
// multiplied by 100000, so dividing back by 100000 until the figure is plausible
// recovers the real rupee value (its leading ~7 significant digits — the rest was
// lost to float precision when it first overflowed, which no amount can recover).
const IMPLAUSIBLE_RUPEES_MIN = 1e10

/**
 * Heal ECV/Contract Amount cells that a repeated Lakhs->rupees migration
 * inflated to absurd magnitudes (e.g. "2.571292e+21"): divide each such cell by
 * 100000 until it lands back in a plausible rupee range. Idempotent — a value
 * already below IMPLAUSIBLE_RUPEES_MIN is left untouched — so it's safe to run on
 * every load. Recovers the leading significant digits; sub-precision lost to the
 * float overflow can't be restored, so the user should spot-check the figure.
 */
export function repairInflatedRupees(table: ExcelTable): ExcelTable {
  const rows = table.rows.map((row) => {
    const next = { ...row }
    for (const col of RUPEE_COLUMNS) {
      const raw = (next[col] ?? '').trim()
      if (!raw) continue
      let n = Number(raw.replace(/,/g, ''))
      if (!Number.isFinite(n) || n < IMPLAUSIBLE_RUPEES_MIN) continue
      while (n >= IMPLAUSIBLE_RUPEES_MIN) n = Math.round(n / 100000)
      next[col] = String(n)
    }
    return next
  })
  return { ...table, rows }
}

/**
 * Like applyWorksSchema, but pulls each standard column's value from
 * whichever imported header `mapping` resolved it to (semantic/keyword
 * match — see core/createDocument.ts's matchPlaceholdersToColumns), not an
 * exact name match. A freshly-downloaded sheet using different column names
 * (e.g. "Estimate Amount" instead of "Amount of estimate") lands in the
 * right place instead of being silently dropped. Falls back to an exact
 * name match for any column the mapping didn't resolve.
 *
 * Any imported column the mapping didn't claim for a standard column (e.g.
 * "Eoffice", "Download start time" — real fields the app doesn't already
 * track) is kept too, appended after the standard columns under its own
 * original name, rather than discarded.
 */
export function applyWorksSchemaWithMapping(
  importedHeaders: string[],
  importedRows: Record<string, string>[],
  mapping: PlaceholderMatch[],
  meta: { id: string; name: string; path: string }
): ExcelTable {
  const columnFor = new Map(mapping.map((m) => [m.label, m.column]))
  const claimed = new Set(mapping.map((m) => m.column).filter((c): c is string => c !== null))
  // A sheet still using "CNO" is folded into "Circle number" rather than kept
  // as a separate extra column (they're the same thing).
  const extraHeaders = importedHeaders.filter((h) => !claimed.has(h) && h !== LEGACY_CNO_COLUMN)
  const headers = [...WORKS_COLUMNS, ...extraHeaders]

  const rows = (importedRows.length > 0 ? importedRows : [{}]).map((row) => {
    const next: Record<string, string> = {}
    for (const h of WORKS_COLUMNS) {
      const matchedCol = columnFor.get(h)
      const legacy = h === 'Circle number' ? row[LEGACY_CNO_COLUMN] : undefined
      next[h] = (matchedCol ? row[matchedCol] : undefined) ?? row[h] ?? legacy ?? ''
    }
    for (const h of extraHeaders) {
      next[h] = row[h] ?? ''
    }
    return next
  })
  return { ...meta, headers, rows }
}

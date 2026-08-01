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
  // Preserve the table's own column arrangement (e.g. the order of the Google
  // sheet it was imported from, so a re-normalize on load never reshuffles it) —
  // rename any legacy column in place, then only *append* standard columns the
  // table is missing. Never reorder what's already there.
  const headers: string[] = []
  const seen = new Set<string>()
  const push = (h: string): void => {
    if (!seen.has(h)) {
      seen.add(h)
      headers.push(h)
    }
  }
  for (const h of table.headers) {
    push(h === LEGACY_ECV_COLUMN ? 'ECV' : h === LEGACY_CNO_COLUMN ? 'Circle number' : h)
  }
  for (const h of WORKS_COLUMNS) push(h)
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
 * track) is kept too, under its own original name, rather than discarded.
 *
 * The output columns follow the *imported sheet's own order*, so a download of
 * the Works List mirrors the arrangement of the Google sheet it came from: each
 * mapped column appears under its standard (app) name at the sheet's position,
 * each unclaimed column keeps its own name in place, and any standard column the
 * sheet didn't carry is appended afterwards so the app never loses a field it
 * relies on.
 */
export function applyWorksSchemaWithMapping(
  importedHeaders: string[],
  importedRows: Record<string, string>[],
  mapping: PlaceholderMatch[],
  meta: { id: string; name: string; path: string }
): ExcelTable {
  const columnFor = new Map(mapping.map((m) => [m.label, m.column]))
  // Reverse of columnFor: which standard column an imported header satisfies.
  const labelForImported = new Map<string, string>()
  for (const m of mapping) if (m.column) labelForImported.set(m.column, m.label)

  const seen = new Set<string>()
  // The sheet's own columns in the sheet's order (mapped to standard names) —
  // exactly what a download reproduces. A "CNO" column folds into "Circle
  // number" (they're the same thing) rather than kept as a separate extra.
  const sourceHeaders: string[] = []
  for (const h of importedHeaders) {
    const name = labelForImported.get(h) ?? (h === LEGACY_CNO_COLUMN ? 'Circle number' : h)
    if (!seen.has(name)) {
      seen.add(name)
      sourceHeaders.push(name)
    }
  }
  // The full column set the app works with: the sheet's columns, then any
  // standard column the sheet didn't include, appended so the app never loses a
  // field it relies on. (Those appended columns aren't part of the download.)
  const headers = [...sourceHeaders]
  for (const h of WORKS_COLUMNS) {
    if (!seen.has(h)) {
      seen.add(h)
      headers.push(h)
    }
  }

  const rows = (importedRows.length > 0 ? importedRows : [{}]).map((row) => {
    const next: Record<string, string> = {}
    for (const h of WORKS_COLUMNS) {
      const matchedCol = columnFor.get(h)
      const legacy = h === 'Circle number' ? row[LEGACY_CNO_COLUMN] : undefined
      next[h] = (matchedCol ? row[matchedCol] : undefined) ?? row[h] ?? legacy ?? ''
    }
    // Any remaining (non-standard) column keeps its own value.
    for (const h of headers) if (!(h in next)) next[h] = row[h] ?? ''
    return next
  })
  return { ...meta, headers, sourceHeaders: sourceHeaders.length > 0 ? sourceHeaders : undefined, rows }
}

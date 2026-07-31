import type { ExcelTable } from '@core/types'
import { splitCircleNumberAndName } from '@core/monitoringImport'
import { resolveFromDirectory, entriesOf, CMC_ZONE_CIRCLES, type ZoneCircleEntry } from './zoneCircleDirectory'

export interface ZoneCircleMismatch {
  rowIndex: number
  workName: string
  foundZone?: string
  foundCircle?: string
}

export interface ZoneCircleResult {
  table: ExcelTable
  /** Rows whose explicit (or name-inferred) Zone/Circle conflicts with the logged-in identity — a non-empty list means the import should be rejected. */
  mismatches: ZoneCircleMismatch[]
  /** How many blank Zone/Circle cells were auto-filled from a match found inside "Name of the work". */
  filledCount: number
}

const norm = (s: string) => s.trim().toLowerCase()

// Same office, allowing for the credentials sheet storing a circle bare
// ("Gajularamaram") while a work name or column spells it out
// ("Gajularamaram Circle-57"): equal once normalized, or one contained in the
// other. No two CMC circle (or zone) names are substrings of each other, so
// this can't conflate two different offices.
const sameId = (a: string, b: string) => {
  const na = norm(a)
  const nb = norm(b)
  if (!na || !nb) return false
  return na === nb || na.includes(nb) || nb.includes(na)
}

/**
 * Enforce that every row of a Works List belongs to the logged-in Head
 * Draughtsman's own Zone/Circle (from the login credentials sheet), filling in
 * the Zone/Circle columns as it goes. The whole list is one Head Draughtsman's
 * own circle, so the default is to stamp every row with the login identity —
 * the checks below only exist to catch a row that plainly belongs to someone
 * else and reject it rather than silently mixing lists:
 * - An explicit Zone/Circle cell that doesn't match the login is a mismatch.
 * - A "Name of the work" that names a *different* CMC circle (an explicit
 *   "<name> circle" tag, or a bare known circle name — see
 *   resolveFromDirectory) is a mismatch, and drags its (foreign) zone with it.
 * - Otherwise the row is the login's own: any blank Zone/Circle is filled with
 *   the login Zone/Circle. Work names here usually mention neither their circle
 *   nor their zone (e.g. "Improvements to Peddamma temple road"), so waiting
 *   for the name to spell them out would leave the columns blank — the reason
 *   they weren't auto-filling before. A conflicting row is never filled (its
 *   import is rejected wholesale by the caller).
 */
export function enforceZoneCircle(
  table: ExcelTable,
  loginZone: string,
  loginCircle: string,
  entries: ZoneCircleEntry[] = CMC_ZONE_CIRCLES
): ZoneCircleResult {
  const zoneHeader = table.headers.find((h) => h.trim().toLowerCase() === 'zone')
  const circleHeader = table.headers.find((h) => h.trim().toLowerCase() === 'circle')
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')

  const mismatches: ZoneCircleMismatch[] = []
  let filledCount = 0

  const rows = table.rows.map((row, rowIndex) => {
    const next = { ...row }
    const workName = (nameHeader ? next[nameHeader] : '') ?? ''
    const nameLower = workName.toLowerCase()
    // Which circle (and hence zone) of the selected corporation the work name
    // points at, if any.
    const dir = resolveFromDirectory(workName, entries)

    const explicitZone = zoneHeader ? (next[zoneHeader] ?? '').trim() : ''
    const explicitCircle = circleHeader ? (next[circleHeader] ?? '').trim() : ''

    // --- Is this row's circle foreign to the login? ---
    let foundCircle: string | undefined
    let circleForeign = false
    if (explicitCircle) {
      if (!sameId(explicitCircle, loginCircle)) {
        foundCircle = explicitCircle
        circleForeign = true
      }
    } else if (loginCircle && nameLower.includes(loginCircle.toLowerCase())) {
      // Name confirms the login's own circle.
    } else if (dir.circle && !sameId(dir.circle, loginCircle)) {
      // Name names a different CMC circle.
      foundCircle = dir.circle
      circleForeign = true
    }

    // --- Is this row's zone foreign to the login? ---
    let foundZone: string | undefined
    if (explicitZone) {
      if (!sameId(explicitZone, loginZone)) foundZone = explicitZone
    } else if (loginZone && nameLower.includes(loginZone.toLowerCase())) {
      // Name confirms the login's own zone.
    } else if (circleForeign && dir.zone && !sameId(dir.zone, loginZone)) {
      // A foreign circle carries its own (foreign) zone.
      foundZone = dir.zone
    }

    if (foundZone || foundCircle) {
      // Belongs to a different office — flag and leave untouched.
      mismatches.push({ rowIndex, workName, foundZone, foundCircle })
      return next
    }

    // The row is the login's own: fill any blank Zone/Circle with the login
    // identity, whether or not the name happened to mention it.
    if (zoneHeader && !explicitZone && loginZone) {
      next[zoneHeader] = loginZone
      filledCount++
    }
    if (circleHeader && !explicitCircle && loginCircle) {
      next[circleHeader] = loginCircle
      filledCount++
    }
    return next
  })

  return { table: { ...table, rows }, mismatches, filledCount }
}

// The Works List's circle-number column, named "Circle number" now but still
// "CNO" on sheets/tables saved before the rename — matched either way.
const findCircleNumberHeader = (table: ExcelTable) =>
  table.headers.find((h) => {
    const n = h.trim().toLowerCase()
    return n === 'circle number' || n === 'cno'
  })

/**
 * Fill any blank Circle number cell with the logged-in Head Draughtsman's own
 * circle number. Unlike Zone/Circle, a circle number isn't something a work's
 * name would ever mention, so there's no name-matching or mismatch check
 * here — just a direct fill of whatever's blank.
 */
export function fillCircleNumber(table: ExcelTable, circleNumber?: string): ExcelTable {
  const cnoHeader = findCircleNumberHeader(table)
  if (!cnoHeader || !circleNumber) return table
  const rows = table.rows.map((row) => {
    if ((row[cnoHeader] ?? '').trim()) return row
    return { ...row, [cnoHeader]: circleNumber }
  })
  return { ...table, rows }
}

/**
 * Pull a circle number embedded in the Circle column out into the Circle
 * number column. A downloaded sheet often writes the Circle cell as a combined
 * "57-Gajularamaram" / "Gajularamaram 57" (either order, hyphen or space) — see
 * splitCircleNumberAndName — so here the bare name stays in Circle and the
 * digits move to Circle number. A Circle cell that carries no number, or a row
 * whose Circle number is already filled, is left untouched (never clobbered).
 * A table without both columns is returned as-is so a number is never dropped.
 */
export function splitCircleColumn(table: ExcelTable): ExcelTable {
  const circleHeader = table.headers.find((h) => h.trim().toLowerCase() === 'circle')
  const cnoHeader = findCircleNumberHeader(table)
  if (!circleHeader || !cnoHeader) return table
  const rows = table.rows.map((row) => {
    const raw = (row[circleHeader] ?? '').trim()
    if (!raw) return row
    const split = splitCircleNumberAndName(raw)
    if (!split) return row
    const next = { ...row, [circleHeader]: split.circle }
    if (!(row[cnoHeader] ?? '').trim()) next[cnoHeader] = split.cno
    return next
  })
  return { ...table, rows }
}

export interface LoginIdentity {
  /** Which corporation's zone/circle directory to validate against (e.g. "CMC"). */
  corporation?: string
  zone?: string
  circle?: string
  circleNumber?: string
}

/**
 * Auto-fill one Works List row's Zone / Circle / Circle number from what's
 * already in it — used live as the user types a work name into the in-app
 * grid, so the derived columns fill in without a re-import. Runs the same
 * three steps an import does, on a single row: split a combined Circle cell,
 * infer Zone/Circle from the "Name of the work" (via the CMC directory), and
 * fill a blank Circle number from the login. Only blank cells are ever
 * written — the user's own entries are never overwritten — and a name that
 * points at a different circle is simply left unfilled here (the import path
 * is where a genuine mismatch is rejected).
 */
export function autofillWorksRow(row: Record<string, string>, login: LoginIdentity): Record<string, string> {
  const one: ExcelTable = { id: '', name: '', path: '', headers: Object.keys(row), rows: [row] }
  let t = splitCircleColumn(one)
  if (login.zone && login.circle) {
    // Fall back to CMC when no corporation is set, preserving prior behaviour.
    const entries = login.corporation ? entriesOf(login.corporation) : CMC_ZONE_CIRCLES
    t = enforceZoneCircle(t, login.zone, login.circle, entries).table
  }
  t = fillCircleNumber(t, login.circleNumber)
  return t.rows[0]
}

import type { ExcelTable } from '@core/types'

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

/**
 * Enforce that every row of an imported Works List belongs to the logged-in
 * Head Draftsman's own Zone/Circle (from the login credentials sheet):
 * - A row with an explicit Zone/Circle that doesn't match is a mismatch.
 * - A row with a blank Zone and/or Circle is checked against "Name of the
 *   work" instead — works are conventionally named mentioning their own
 *   Zone/Circle — and if the logged-in Zone/Circle appears there, the blank
 *   cell is filled in automatically.
 * - A row where neither an explicit value nor one inferable from the name
 *   exists is left alone — there's nothing to check it against, so it isn't
 *   treated as a mismatch.
 */
export function enforceZoneCircle(table: ExcelTable, loginZone: string, loginCircle: string): ZoneCircleResult {
  const zoneHeader = table.headers.find((h) => h.trim().toLowerCase() === 'zone')
  const circleHeader = table.headers.find((h) => h.trim().toLowerCase() === 'circle')
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')

  const mismatches: ZoneCircleMismatch[] = []
  let filledCount = 0

  const rows = table.rows.map((row, rowIndex) => {
    const next = { ...row }
    const workName = (nameHeader ? next[nameHeader] : '') ?? ''
    let foundZone: string | undefined
    let foundCircle: string | undefined

    if (zoneHeader) {
      const explicit = (next[zoneHeader] ?? '').trim()
      if (explicit) {
        if (norm(explicit) !== norm(loginZone)) foundZone = explicit
      } else if (workName && loginZone && workName.toLowerCase().includes(loginZone.toLowerCase())) {
        next[zoneHeader] = loginZone
        filledCount++
      }
    }

    if (circleHeader) {
      const explicit = (next[circleHeader] ?? '').trim()
      if (explicit) {
        if (norm(explicit) !== norm(loginCircle)) foundCircle = explicit
      } else if (workName && loginCircle && workName.toLowerCase().includes(loginCircle.toLowerCase())) {
        next[circleHeader] = loginCircle
        filledCount++
      }
    }

    if (foundZone || foundCircle) mismatches.push({ rowIndex, workName, foundZone, foundCircle })
    return next
  })

  return { table: { ...table, rows }, mismatches, filledCount }
}

/**
 * Fill any blank CNO (Circle number) cell with the logged-in Head Draftsman's
 * own circle number. Unlike Zone/Circle, a circle number isn't something a
 * work's name would ever mention, so there's no name-matching or mismatch
 * check here — just a direct fill of whatever's blank.
 */
export function fillCircleNumber(table: ExcelTable, circleNumber?: string): ExcelTable {
  const cnoHeader = table.headers.find((h) => h.trim().toLowerCase() === 'cno')
  if (!cnoHeader || !circleNumber) return table
  const rows = table.rows.map((row) => {
    if ((row[cnoHeader] ?? '').trim()) return row
    return { ...row, [cnoHeader]: circleNumber }
  })
  return { ...table, rows }
}

import type { ExcelTable } from './types'
import { normalizeWorkNameForMatch as norm, reservationCategory, reservationCategoryFromRow } from './tenderAgents/nameOfWork'

/**
 * Rules for keeping the Works List database (the source of truth — L1/LOA scans keep updating it) in step with the
 * Google Sheet it was first imported from.
 */

const header = (table: ExcelTable, name: string) => table.headers.find((h) => h.trim().toLowerCase() === name)

const cell = (row: Record<string, string>, h: string | undefined) => (h ? (row[h] ?? '').trim() : '')

/** True once the database holds at least one real work (the starter table's single blank row doesn't count). */
export function hasWorks(table: ExcelTable | undefined): boolean {
  if (!table) return false
  const nameH = header(table, 'name of the work')
  const codeH = header(table, 'wincode')
  return table.rows.some((r) => cell(r, nameH) !== '' || cell(r, codeH) !== '')
}

/** The only fields a refresh may change on a work that is already in the database (the sheet owns these). */
const SHEET_OWNED = ['wincode', 'eoffice', 'amount of estimate'] as const

/**
 * Re-import / refresh: append the sheet's NEW works and leave every existing work as it is — except the sheet-owned
 * fields (Wincode, Eoffice number, Amount of estimate), which follow the sheet when it carries a different non-blank
 * value. A work is "existing" when its Wincode, or (failing that) its normalized name, is already in the database;
 * everything else on it (ECV, tender details, ...) comes from scans and is never overwritten.
 */
export function mergeNewWorks(
  existing: ExcelTable,
  imported: ExcelTable
): { table: ExcelTable; added: number; skipped: number; updated: number } {
  const exNameH = header(existing, 'name of the work')
  const exCodeH = header(existing, 'wincode')
  const imNameH = header(imported, 'name of the work')
  const imCodeH = header(imported, 'wincode')

  const headers = [...existing.headers]
  // Sheet-owned columns present in the sheet: [column in the sheet, column in the database (added if missing)].
  const owned: [string, string][] = []
  for (const key of SHEET_OWNED) {
    const imH = header(imported, key)
    if (!imH) continue
    let exH = header(existing, key)
    if (!exH) {
      exH = imH
      headers.push(exH)
    }
    owned.push([imH, exH])
  }

  const rows = existing.rows.filter((r) => Object.values(r).some((v) => (v ?? '').trim() !== '')).map((r) => ({ ...r }))
  const byCode = new Map<string, number>()
  const byName = new Map<string, number>()
  rows.forEach((r, i) => {
    const code = cell(r, exCodeH).toLowerCase()
    const name = norm(cell(r, exNameH))
    if (code && !byCode.has(code)) byCode.set(code, i)
    if (name && !byName.has(name)) byName.set(name, i)
  })

  let added = 0
  let skipped = 0
  let updated = 0
  for (const r of imported.rows) {
    const code = cell(r, imCodeH).toLowerCase()
    const name = norm(cell(r, imNameH))
    if (!code && !name) continue
    const at = (code ? byCode.get(code) : undefined) ?? (name ? byName.get(name) : undefined)
    if (at != null) {
      skipped++
      let changed = false
      for (const [imH, exH] of owned) {
        const v = (r[imH] ?? '').trim()
        if (v && v !== (rows[at][exH] ?? '').trim()) {
          rows[at][exH] = v
          changed = true
        }
      }
      if (changed) updated++
      continue
    }
    const row: Record<string, string> = {}
    for (const h of headers) {
      const imH = imported.headers.includes(h) ? h : owned.find(([, e]) => e === h)?.[0]
      row[h] = imH ? (r[imH] ?? '') : ''
    }
    rows.push(row)
    if (code) byCode.set(code, rows.length - 1)
    if (name) byName.set(name, rows.length - 1)
    added++
  }

  return { table: fillReservations({ ...existing, headers, rows }), added, skipped, updated }
}

/** True when this work name is already a row of the Works List (exact after normalizing, or one contains the other). */
export function workInDatabase(table: ExcelTable | undefined, workName: string | undefined): boolean {
  const target = norm(workName ?? '')
  if (!table || !target) return false
  const nameH = header(table, 'name of the work')
  return table.rows.some((r) => {
    const n = norm(cell(r, nameH))
    return (
      !!n && (n === target || (Math.min(n.length, target.length) >= 25 && (n.includes(target) || target.includes(n))))
    )
  })
}

const sameCircle = (a: string, b: string) => {
  const na = a.trim().toLowerCase()
  const nb = b.trim().toLowerCase()
  return !!na && !!nb && (na === nb || na.includes(nb) || nb.includes(na))
}

/**
 * Does a scanned L1 sheet / LOA belong to the selected office? `sheetCircle` is the circle the sheet itself points at
 * (from its NIT No or the work name). A sheet naming another circle is rejected — unless its work is already in this
 * office's Works List: GHMC was split into CMC, MMC and GHMC in Feb 2026, so a pre-split work can still be this
 * office's while its name names the old circle. A sheet that names no circle at all, or an office with no circle
 * (Zone level), is let through.
 */
export function scanBelongsToOffice(opts: {
  officeCircle: string | undefined
  sheetCircle: string | undefined
  workName: string | undefined
  table: ExcelTable | undefined
}): { ok: true } | { ok: false; error: string } {
  const office = (opts.officeCircle ?? '').trim()
  const sheet = (opts.sheetCircle ?? '').trim()
  if (!office || !sheet || sameCircle(office, sheet)) return { ok: true }
  if (workInDatabase(opts.table, opts.workName)) return { ok: true }
  return { ok: false, error: `These sheets do not belong to the ${office} office.` }
}

/** Works List columns the Agreement & Work Order workspace's own inputs are saved into. */
export const WORKSPACE_DB_COLUMNS = {
  adminSanctionDate: 'Admin Sanction Date',
  newspapers: 'Newspapers Published',
  aeNameNumber: 'AE Name & Number',
  adminSanctionBy: 'Sanction By'
} as const

export type WorkspaceInputs = Partial<Record<keyof typeof WORKSPACE_DB_COLUMNS, string>>

/** Writes the workspace's inputs onto one Works List row. Returns the same table object when nothing changed. */
export function saveWorkspaceInputs(table: ExcelTable, rowIndex: number, values: WorkspaceInputs): ExcelTable {
  const row = table.rows[rowIndex]
  if (!row) return table
  const next = { ...row }
  let changed = false
  for (const [k, v] of Object.entries(values) as [keyof typeof WORKSPACE_DB_COLUMNS, string][]) {
    const col = WORKSPACE_DB_COLUMNS[k]
    if (!table.headers.includes(col)) continue
    if ((next[col] ?? '') !== v) {
      next[col] = v
      changed = true
    }
  }
  if (!changed) return table
  return { ...table, rows: table.rows.map((r, i) => (i === rowIndex ? next : r)) }
}

/** What the Reservation column reads for a work: the category named in its title (SC, ST, Waddera, ...), else "General". */
export function reservationLabel(workName: string): string {
  return reservationCategory(workName) || 'General'
}

/**
 * Writes each work's actual reservation into the Reservation column — the category its name carries ("SC", "ST",
 * "Waddera", ...) or "General" when it names none. A category typed by hand against a name that carries none is kept.
 * Rows with no work name are left alone. Returns the same table object when nothing changed.
 */
export function fillReservations(table: ExcelTable): ExcelTable {
  const nameH = header(table, 'name of the work')
  const resH = header(table, 'reservation')
  if (!nameH || !resH) return table
  let changed = false
  const rows = table.rows.map((r) => {
    if (!cell(r, nameH)) return r
    const want = reservationCategoryFromRow({ 'Name of the work': r[nameH] ?? '', Reservation: r[resH] ?? '' }) || 'General'
    if ((r[resH] ?? '') === want) return r
    changed = true
    return { ...r, [resH]: want }
  })
  return changed ? { ...table, rows } : table
}

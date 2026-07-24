import type { ExcelTable } from './types'
import type { SheetGrid } from './sheet'
import type { PlaceholderMatch } from './createDocument'

const norm = (s: string) => s.trim().toLowerCase()

function tokens(s: string): string[] {
  return norm(s)
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
}

/**
 * Pick the one sheet whose name matches the logged-in Circle. A monitoring
 * workbook holds one sheet per circle (grouped under zone tabs) — an exact
 * normalized match wins; failing that, a token-overlap match handles a sheet
 * name that abbreviates or reorders the circle's full name (e.g.
 * "Gajularamaram-57" vs the login's "Gajularamaram Circle-57" — a plain
 * substring check misses this since "Circle" sits in between). A sheet only
 * qualifies when every token of the shorter of the two names appears in the
 * other — better to find nothing than confidently pick the wrong circle.
 */
export function findCircleSheet(sheets: SheetGrid[], circle: string): SheetGrid | null {
  const target = norm(circle)
  if (!target) return null
  const exact = sheets.find((s) => norm(s.sheetName) === target)
  if (exact) return exact

  const targetTokens = tokens(circle)
  if (targetTokens.length === 0) return null
  let best: SheetGrid | null = null
  let bestScore = 0
  for (const s of sheets) {
    const sheetTokens = tokens(s.sheetName)
    if (sheetTokens.length === 0) continue
    const shared = sheetTokens.filter((t) => targetTokens.includes(t)).length
    const score = shared / Math.min(sheetTokens.length, targetTokens.length)
    if (score > bestScore) {
      bestScore = score
      best = s
    }
  }
  return bestScore >= 1 ? best : null
}

export interface MonitoringMergeResult {
  table: ExcelTable
  /** Monitoring rows matched to an existing Works List row by Wincode or Name of the work. */
  matchedCount: number
  /** Monitoring rows matching nothing existing, added as new rows. */
  addedCount: number
  /** Blank cells filled in on matched rows. Existing (non-blank) values are never overwritten. */
  filledCount: number
}

/**
 * Merge a monitoring-format circle sheet into the existing Works List. Each
 * monitoring row is matched to an existing row by Wincode (preferred) or
 * Name of the work (fallback):
 * - A match only has its *blank* cells filled in — existing data is never
 *   overwritten by the monitoring sheet.
 * - A monitoring row matching nothing existing becomes a new row.
 *
 * `mapping` resolves each Works List column to whichever monitoring header
 * means the same thing — a monitoring workbook won't use the app's exact
 * column names, so this reuses the same semantic/keyword matching as
 * document placeholders (core/createDocument.ts's matchPlaceholdersToColumns,
 * called with labels = Works List columns, columns = monitoring headers).
 */
export function mergeMonitoringRows(
  table: ExcelTable,
  monitoringRows: Record<string, string>[],
  mapping: PlaceholderMatch[]
): MonitoringMergeResult {
  const columnFor = new Map(mapping.map((m) => [m.label, m.column]))
  const wincodeMonCol = columnFor.get('Wincode') ?? null
  const nameMonCol = columnFor.get('Name of the work') ?? null

  const rows = table.rows.map((r) => ({ ...r }))
  let matchedCount = 0
  let filledCount = 0
  const newRows: Record<string, string>[] = []

  for (const monRow of monitoringRows) {
    const monWincode = wincodeMonCol ? (monRow[wincodeMonCol] ?? '').trim() : ''
    const monName = nameMonCol ? (monRow[nameMonCol] ?? '').trim() : ''
    if (!monWincode && !monName) continue // nothing to match or usefully add on

    let matchIndex = -1
    if (monWincode) {
      matchIndex = rows.findIndex((r) => norm(r['Wincode'] ?? '') === norm(monWincode))
    }
    if (matchIndex === -1 && monName) {
      matchIndex = rows.findIndex((r) => norm(r['Name of the work'] ?? '') === norm(monName))
    }

    if (matchIndex === -1) {
      const newRow: Record<string, string> = {}
      for (const h of table.headers) {
        const monCol = columnFor.get(h)
        newRow[h] = monCol ? (monRow[monCol] ?? '') : ''
      }
      newRows.push(newRow)
      continue
    }

    matchedCount++
    const existing = rows[matchIndex]
    for (const h of table.headers) {
      if ((existing[h] ?? '').trim()) continue
      const monCol = columnFor.get(h)
      const value = monCol ? (monRow[monCol] ?? '').trim() : ''
      if (value) {
        existing[h] = value
        filledCount++
      }
    }
  }

  return {
    table: { ...table, rows: [...rows, ...newRows] },
    matchedCount,
    addedCount: newRows.length,
    filledCount
  }
}

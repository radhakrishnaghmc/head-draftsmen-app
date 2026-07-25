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
 * Extracts a bare circle number from a sheet name — this app's monitoring
 * workbooks name their per-circle tabs after the circle's number rather
 * than its full name, most often as "c57"/"C-58"/"Circle 59" (hardcoded
 * here since that's this specific naming convention, not a general rule),
 * or occasionally just the bare number on its own ("58"). Returns null for
 * a sheet name that names its circle no other way (a summary/index tab, a
 * sheet already named after the circle's full name, etc).
 */
function circleNumberFromSheetName(name: string): string | null {
  const withPrefix = /\bc(?:ircle)?[\s.\-_#]*?(\d{1,4})\b/i.exec(name)
  if (withPrefix) return withPrefix[1]
  const bare = /^\s*(\d{1,4})\s*$/.exec(name)
  return bare ? bare[1] : null
}

/**
 * Pick the one sheet whose name matches the logged-in Circle. A monitoring
 * workbook holds one sheet per circle (grouped under zone tabs), most often
 * named after the circle's *number* (see circleNumberFromSheetName) rather
 * than its full name — tried first when `circleNumber` is available, since
 * it's this app's actual real-world sheet-naming convention. Falls back to
 * matching the sheet name against the circle's own full name: an exact
 * normalized match, then a token-overlap match for a sheet name that
 * abbreviates or reorders it (e.g. "Gajularamaram-57" vs the login's
 * "Gajularamaram Circle-57" — a plain substring check misses this since
 * "Circle" sits in between). A sheet only qualifies on token overlap when
 * every token of the shorter of the two names appears in the other —
 * better to find nothing than confidently pick the wrong circle.
 *
 * A real monitoring workbook pairs each circle's number with *two* sheets —
 * e.g. "C57 MF" and "C57 list of works" — and only one of them, "list of
 * works", actually lists individual works with columns like Name of the
 * Work/Wincode/Agency; the "MF" ("Monitoring Format") sheet is an
 * aggregated pivot/dashboard (category totals, no per-work rows at all) and
 * would map onto nothing useful if merged. When more than one sheet shares
 * the target circle number, whichever one's name says "list of works" wins
 * outright over a same-numbered summary sheet.
 */
export function findCircleSheet(sheets: SheetGrid[], circle: string, circleNumber?: string): SheetGrid | null {
  const targetNumber = (circleNumber ?? '').replace(/\D/g, '')
  if (targetNumber) {
    const candidates = sheets.filter((s) => circleNumberFromSheetName(s.sheetName) === targetNumber)
    if (candidates.length > 0) {
      const listOfWorks = candidates.find((s) => /list\s+of\s+works/i.test(s.sheetName))
      return listOfWorks ?? candidates[0]
    }
  }

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

export interface AgencySplit {
  name: string
  /** 10-digit phone numbers found embedded in the raw agency name text, in the order they appeared. */
  phones: string[]
}

/**
 * A monitoring sheet's "Agency" column often carries the contractor's phone
 * number(s) run straight into the name text (e.g. "Radha Krishna Contractors
 * 9789879878") rather than in a column of their own. Splits any whitespace-
 * delimited token that's exactly 10 digits (once its own punctuation —
 * hyphens, dots — is stripped) out as a phone number, leaving the rest
 * joined back together as the clean name. A token with a different digit
 * count (a registration number, a PIN code, ...) is left in the name as-is,
 * since 10 digits is the one reliable signal an embedded phone number gives
 * — better to miss an oddly-formatted number than misidentify unrelated
 * digits as one.
 */
export function splitAgencyNameAndPhones(raw: string): AgencySplit {
  const nameTokens: string[] = []
  const phones: string[] = []
  for (const tok of raw.trim().split(/\s+/)) {
    const digits = tok.replace(/\D/g, '')
    if (digits.length === 10) {
      phones.push(digits)
    } else if (tok) {
      nameTokens.push(tok)
    }
  }
  return { name: nameTokens.join(' ').replace(/[,;]+$/, '').trim(), phones }
}

/** "1. 9789879878\n2. 9848012345" for more than one number; the bare number for exactly one; "" for none. */
export function formatAgencyPhones(phones: string[]): string {
  if (phones.length <= 1) return phones[0] ?? ''
  return phones.map((p, i) => `${i + 1}. ${p}`).join('\n')
}

export interface CircleSplit {
  cno: string
  circle: string
}

/**
 * A monitoring sheet's own "Circle" column is commonly written as the
 * circle's number and name glued together with a hyphen — "58-Nizampet" —
 * the same "two distinct fields squeezed into one cell" shape as the
 * agency name/phone problem above, just for Circle/CNO instead of Name of
 * the Agency/Phone number of the agency. Returns null for a value that
 * isn't in that "<digits>-<name>" shape (a sheet that already keeps them
 * separate, or one with just a bare circle name) — never guessed at.
 */
export function splitCircleNumberAndName(raw: string): CircleSplit | null {
  const m = /^\s*(\d+)\s*-\s*(.+?)\s*$/.exec(raw)
  return m ? { cno: m[1], circle: m[2] } : null
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
 * "Name of the Agency" is special-cased: whatever monitoring column maps to
 * it is run through splitAgencyNameAndPhones first, so an embedded phone
 * number lands in "Phone number of the agency" instead of staying stuck in
 * the name — only when that split actually finds a 10-digit number, so a
 * monitoring sheet with a genuinely separate phone column (mapped normally,
 * like every other column) isn't overridden by an empty result.
 *
 * "Circle" is special-cased the same way: whatever monitoring column maps
 * to it is run through splitCircleNumberAndName first, so a "58-Nizampet"
 * style combined value fills CNO and Circle separately instead of landing
 * the whole "58-Nizampet" string in Circle alone and leaving CNO blank —
 * only when that split actually recognizes the "<digits>-<name>" shape, so
 * a monitoring sheet with a genuinely separate CNO column isn't overridden.
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
  const agencyMonCol = columnFor.get('Name of the Agency') ?? null
  const circleMonCol = columnFor.get('Circle') ?? null

  function agencySplitFor(monRow: Record<string, string>): AgencySplit | null {
    if (!agencyMonCol) return null
    const raw = (monRow[agencyMonCol] ?? '').trim()
    return raw ? splitAgencyNameAndPhones(raw) : null
  }

  function circleSplitFor(monRow: Record<string, string>): CircleSplit | null {
    if (!circleMonCol) return null
    return splitCircleNumberAndName((monRow[circleMonCol] ?? '').trim())
  }

  function valueFor(
    h: string,
    monRow: Record<string, string>,
    agencySplit: AgencySplit | null,
    circleSplit: CircleSplit | null
  ): string {
    if (h === 'Name of the Agency' && agencySplit) return agencySplit.name
    if (h === 'Phone number of the agency' && agencySplit?.phones.length) {
      return formatAgencyPhones(agencySplit.phones)
    }
    if (h === 'Circle' && circleSplit) return circleSplit.circle
    if (h === 'CNO' && circleSplit) return circleSplit.cno
    const monCol = columnFor.get(h)
    return monCol ? (monRow[monCol] ?? '') : ''
  }

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

    const agencySplit = agencySplitFor(monRow)
    const circleSplit = circleSplitFor(monRow)

    if (matchIndex === -1) {
      const newRow: Record<string, string> = {}
      for (const h of table.headers) newRow[h] = valueFor(h, monRow, agencySplit, circleSplit).trim()
      newRows.push(newRow)
      continue
    }

    matchedCount++
    const existing = rows[matchIndex]
    for (const h of table.headers) {
      if ((existing[h] ?? '').trim()) continue
      const value = valueFor(h, monRow, agencySplit, circleSplit).trim()
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

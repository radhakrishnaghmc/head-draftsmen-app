// The Works List agent — the single place that knows how to audit, query,
// and (eventually) update the app's central Works List database. Every
// document filler / tool that needs to look something up in the Works List
// should route the query through here rather than filtering table.rows
// inline itself — this file is where "is the Works List itself correct" and
// "which row does this query mean" questions get answered, once, rather
// than reimplemented per caller.
import type { ExcelTable } from './types'
import { normalizeWorkNameForMatch } from './tenderAgents/nameOfWork'
import { lakhsToRupees, rupeesFromCell } from './worksAmounts'
import type { MonitoringFormatSummary, MonitoringFormatWorkRow, MonitoringFormatRow } from './monitoringFormat'

export interface WorksListViolation {
  type:
    | 'duplicate-wincode'
    | 'duplicate-name'
    | 'duplicate-tender-id'
    | 'ecv-exceeds-estimate'
    | 'emd-inconsistent'
    | 'mf-abstract-mismatch'
  /** The Wincode/Tender ID/work name/item-type/row this violation centres on — meaning depends on `type`. */
  key: string
  /** Every row index (0-based, into table.rows) involved — empty for a Monitoring Format mismatch, which isn't tied to a Works List row. */
  rowIndices: number[]
  /** A human-readable explanation for a UI banner/list. */
  message: string
}

/**
 * Checks the Works List's core identity invariant: a Wincode identifies
 * exactly one work, and a work has exactly one Wincode (point the user
 * raised directly — two rows quietly sharing a Wincode, or one work typed in
 * twice under different Wincodes, both silently break every Wincode-keyed
 * lookup and update in the app).
 *
 * Blank Wincodes/names are never flagged — a work not yet allotted a Wincode,
 * or an incomplete row, isn't a violation on its own, only two or more
 * ACTUAL values that collide. Work names are compared via
 * normalizeWorkNameForMatch (core/tenderAgents/nameOfWork.ts), which strips
 * reservation/recall/item-number tags first, so a work correctly re-entered
 * under the SAME Wincode with a "(2nd Recall)" tag on a later row isn't a
 * false "duplicate-wincode" hit — only a genuinely different work sharing a
 * Wincode counts.
 */
export function findWincodeUniquenessViolations(table: ExcelTable): WorksListViolation[] {
  const violations: WorksListViolation[] = []
  const byWincode = new Map<string, { rowIndex: number; name: string }[]>()
  const byName = new Map<string, { rowIndex: number; wincode: string; rawName: string }[]>()

  table.rows.forEach((row, rowIndex) => {
    const wincode = (row['Wincode'] ?? '').trim()
    const rawName = row['Name of the work'] ?? ''
    const name = normalizeWorkNameForMatch(rawName)
    if (wincode) {
      const bucket = byWincode.get(wincode) ?? []
      bucket.push({ rowIndex, name })
      byWincode.set(wincode, bucket)
    }
    if (name) {
      const bucket = byName.get(name) ?? []
      bucket.push({ rowIndex, wincode, rawName: rawName.trim() })
      byName.set(name, bucket)
    }
  })

  for (const [wincode, entries] of byWincode) {
    const distinctNames = new Set(entries.map((e) => e.name).filter(Boolean))
    if (distinctNames.size > 1) {
      violations.push({
        type: 'duplicate-wincode',
        key: wincode,
        rowIndices: entries.map((e) => e.rowIndex),
        message: `Wincode "${wincode}" is used for ${distinctNames.size} different works — a Wincode should identify exactly one work.`
      })
    }
  }

  for (const [, entries] of byName) {
    const distinctWincodes = new Set(entries.map((e) => e.wincode).filter(Boolean))
    if (distinctWincodes.size > 1) {
      violations.push({
        type: 'duplicate-name',
        key: entries[0].rawName,
        rowIndices: entries.map((e) => e.rowIndex),
        message: `"${entries[0].rawName}" is entered under ${distinctWincodes.size} different Wincodes (${[...distinctWincodes].join(', ')}) — it should have just one.`
      })
    }
  }

  return violations
}

/**
 * Tender ID identity: a Tender ID names exactly one work — but, unlike
 * Wincode, NOT the reverse. When a tender is recalled after no (responsive)
 * bidder participates, the SAME work is re-tendered under a brand new
 * Tender ID, so one work legitimately carries several Tender IDs over its
 * history; only two DIFFERENT works sharing one Tender ID is an error.
 *
 * NIT No is deliberately never checked for uniqueness at all, in either
 * direction — a single NIT can legitimately bundle several different works
 * as separate items, so many rows sharing one NIT No is completely normal
 * and not worth flagging.
 */
export function findTenderIdUniquenessViolations(table: ExcelTable): WorksListViolation[] {
  const violations: WorksListViolation[] = []
  const byTenderId = new Map<string, { rowIndex: number; name: string }[]>()

  table.rows.forEach((row, rowIndex) => {
    const tenderId = (row['Tender ID'] ?? '').trim()
    if (!tenderId) return
    const name = normalizeWorkNameForMatch(row['Name of the work'] ?? '')
    const bucket = byTenderId.get(tenderId) ?? []
    bucket.push({ rowIndex, name })
    byTenderId.set(tenderId, bucket)
  })

  for (const [tenderId, entries] of byTenderId) {
    const distinctNames = new Set(entries.map((e) => e.name).filter(Boolean))
    if (distinctNames.size > 1) {
      violations.push({
        type: 'duplicate-tender-id',
        key: tenderId,
        rowIndices: entries.map((e) => e.rowIndex),
        message: `Tender ID "${tenderId}" is used for ${distinctNames.size} different works — a Tender ID should identify exactly one work (the same work CAN legitimately carry several Tender IDs across recalls, just not the other way round).`
      })
    }
  }

  return violations
}

/** Indian-grouped rupee figure for a violation message, without pulling in the full formatRupees "Rs …/-" wrapper. */
function rupeeFigure(n: number): string {
  return n.toLocaleString('en-IN')
}

/**
 * ECV can never exceed the Amount of estimate — the tender is invited FOR
 * that estimate, so its contract value (ECV) is by definition capped at it.
 * A row where ECV reads higher almost always means a data-entry mistake
 * (e.g. the estimate cell holding a stale/wrong figure, or ECV entered in
 * the wrong unit). Only checked when both cells are actually filled in.
 */
export function findEcvExceedsEstimateViolations(table: ExcelTable): WorksListViolation[] {
  const violations: WorksListViolation[] = []
  table.rows.forEach((row, rowIndex) => {
    const estimateRaw = (row['Amount of estimate'] ?? '').trim()
    const ecvRaw = (row['ECV'] ?? '').trim()
    if (!estimateRaw || !ecvRaw) return
    const estimate = lakhsToRupees(estimateRaw)
    const ecv = rupeesFromCell(ecvRaw)
    if (ecv > estimate) {
      violations.push({
        type: 'ecv-exceeds-estimate',
        key: `row-${rowIndex}`,
        rowIndices: [rowIndex],
        message: `Row ${rowIndex + 1}: ECV (Rs ${rupeeFigure(ecv)}) is greater than Amount of estimate (Rs ${rupeeFigure(estimate)}) — ECV should never exceed the estimate.`
      })
    }
  })
  return violations
}

/**
 * EMD @ 1.5% of ECV is mathematically always greater than EMD @ 1% of the
 * same ECV — a row where the stored EMD 1.5% cell reads less than or equal
 * to EMD 1% means one of the two was entered/computed wrong (transposed
 * figures, a stale value left over from an earlier ECV, …). Checked on the
 * stored cell values themselves (not recomputed from ECV), since that's
 * exactly what could be inconsistent. Only checked when both are filled in.
 */
export function findEmdInconsistencyViolations(table: ExcelTable): WorksListViolation[] {
  const violations: WorksListViolation[] = []
  table.rows.forEach((row, rowIndex) => {
    const emd1Raw = (row['EMD 1%'] ?? '').trim()
    const emd15Raw = (row['EMD 1.5%'] ?? '').trim()
    if (!emd1Raw || !emd15Raw) return
    const emd1 = rupeesFromCell(emd1Raw)
    const emd15 = rupeesFromCell(emd15Raw)
    if (emd15 <= emd1) {
      violations.push({
        type: 'emd-inconsistent',
        key: `row-${rowIndex}`,
        rowIndices: [rowIndex],
        message: `Row ${rowIndex + 1}: EMD 1.5% (Rs ${rupeeFigure(emd15)}) should be greater than EMD 1% (Rs ${rupeeFigure(emd1)}), not less than or equal to it.`
      })
    }
  })
  return violations
}

/** Abstract row label, normalized for matching between the Abstract and the list of works — same source workbook, but casing/spacing sometimes differs between the two sheets. */
function normalizeAbstractLabel(s: string): string {
  return s.trim().toLowerCase().replace(/\s+/g, ' ')
}

/**
 * Cross-checks one Abstract pivot's per-group totals (per item type, or per
 * ward) against the workbook's own "list of works" sheets — the Abstract is
 * meant to be a roll-up of exactly those rows, so if one sheet gets
 * hand-edited (or a row added/removed) without the other following, the two
 * silently drift apart. Group labels are matched case/whitespace-insensitively
 * since the same label is often typed slightly differently on the Abstract
 * vs. the list. `groupOf` picks the list-of-works column that corresponds to
 * the Abstract rows being checked (typeOfWork for the item-type pivot, ward
 * for the ward-wise one).
 */
function findAbstractTallyMismatches(
  abstractRows: MonitoringFormatRow[],
  works: MonitoringFormatWorkRow[],
  groupOf: (w: MonitoringFormatWorkRow) => string
): WorksListViolation[] {
  const violations: WorksListViolation[] = []
  const byGroup = new Map<string, MonitoringFormatWorkRow[]>()
  for (const w of works) {
    const key = normalizeAbstractLabel(groupOf(w))
    if (!key) continue
    const bucket = byGroup.get(key) ?? []
    bucket.push(w)
    byGroup.set(key, bucket)
  }

  for (const row of abstractRows) {
    const key = normalizeAbstractLabel(row.itemType)
    if (!key || key === 'total') continue
    const matches = byGroup.get(key) ?? []
    const actualNo = matches.length
    const actualAmt = matches.reduce((sum, w) => sum + w.estimateAmt, 0)
    const noMismatch = actualNo !== row.totalWorks.no
    // Amounts come off a spreadsheet, so allow for float/rounding noise rather
    // than flagging a paise-level difference that isn't a real data problem.
    const amtMismatch = Math.abs(actualAmt - row.totalWorks.amt) >= 1
    if (!noMismatch && !amtMismatch) continue
    const parts: string[] = []
    if (noMismatch) {
      parts.push(`${row.totalWorks.no} work${row.totalWorks.no === 1 ? '' : 's'} in the Abstract vs ${actualNo} in the list of works`)
    }
    if (amtMismatch) {
      parts.push(`Rs ${rupeeFigure(row.totalWorks.amt)} in the Abstract vs Rs ${rupeeFigure(actualAmt)} in the list of works`)
    }
    violations.push({
      type: 'mf-abstract-mismatch',
      key: row.itemType,
      rowIndices: [],
      message: `"${row.itemType}": ${parts.join(' and ')} — Abstract and list of works don't tally.`
    })
  }
  return violations
}

/**
 * Every Abstract-vs-list-of-works tally check: the per-item-type pivot
 * (always present), plus the per-ward pivot when the sheet has one.
 */
export function findMonitoringFormatMismatches(
  summary: MonitoringFormatSummary,
  works: MonitoringFormatWorkRow[]
): WorksListViolation[] {
  const violations = findAbstractTallyMismatches(summary.rows, works, (w) => w.typeOfWork)
  if (summary.wardRows) {
    violations.push(...findAbstractTallyMismatches(summary.wardRows, works, (w) => w.ward))
  }
  return violations
}

/**
 * The same Wincode/name identity rules checked on the main Works List
 * (findWincodeUniquenessViolations) apply just as much within the
 * Monitoring Format's own "list of works" rows — Tender ID, ECV, and EMD
 * aren't columns the Monitoring Format carries at all, so only the identity
 * checks (not the amount ones) carry over here.
 */
export function findMonitoringFormatWorksIdentityViolations(works: MonitoringFormatWorkRow[]): WorksListViolation[] {
  const violations: WorksListViolation[] = []
  const byWincode = new Map<string, { slNo: string; name: string }[]>()
  const byName = new Map<string, { slNo: string; wincode: string; rawName: string }[]>()

  works.forEach((w) => {
    const wincode = w.winCode.trim()
    const rawName = w.workName
    const name = normalizeWorkNameForMatch(rawName)
    if (wincode) {
      const bucket = byWincode.get(wincode) ?? []
      bucket.push({ slNo: w.slNo, name })
      byWincode.set(wincode, bucket)
    }
    if (name) {
      const bucket = byName.get(name) ?? []
      bucket.push({ slNo: w.slNo, wincode, rawName: rawName.trim() })
      byName.set(name, bucket)
    }
  })

  for (const [wincode, entries] of byWincode) {
    const distinctNames = new Set(entries.map((e) => e.name).filter(Boolean))
    if (distinctNames.size > 1) {
      violations.push({
        type: 'duplicate-wincode',
        key: `mf-${wincode}`,
        rowIndices: [],
        message: `List of works: Wincode "${wincode}" is used for ${distinctNames.size} different works (Sl.No ${entries.map((e) => e.slNo).join(', ')}) — a Wincode should identify exactly one work.`
      })
    }
  }

  for (const [, entries] of byName) {
    const distinctWincodes = new Set(entries.map((e) => e.wincode).filter(Boolean))
    if (distinctWincodes.size > 1) {
      violations.push({
        type: 'duplicate-name',
        key: `mf-${entries[0].rawName}`,
        rowIndices: [],
        message: `List of works: "${entries[0].rawName}" is entered under ${distinctWincodes.size} different Wincodes (${[...distinctWincodes].join(', ')}) — it should have just one.`
      })
    }
  }

  return violations
}

/** Every Monitoring Format check in one call (item-type tally + the list of works' own identity rules) — the entry point the Errors button (App.tsx) uses alongside findWorksListErrors. */
export function findMonitoringFormatErrors(
  summary: MonitoringFormatSummary,
  works: MonitoringFormatWorkRow[]
): WorksListViolation[] {
  return [...findMonitoringFormatMismatches(summary, works), ...findMonitoringFormatWorksIdentityViolations(works)]
}

export interface WorksListSearchResult {
  row: Record<string, string>
  rowIndex: number
}

/**
 * Substring-searches the Works List by Name of Work — the query behind
 * every "type to find a work" search box (Issue Documents' work picker, and
 * any future one), so that lookup logic lives in one place instead of each
 * caller filtering table.rows inline itself. Case-insensitive; falls back to
 * the table's first column if no header looks like "Name of (the) Work"
 * (matching how the Issue Documents picker itself used to fail-soft).
 */
export function searchWorksList(table: ExcelTable, query: string): WorksListSearchResult[] {
  const nameHeader = table.headers.find((h) => /name of (the )?work/i.test(h)) ?? table.headers[0]
  const all = table.rows.map((row, rowIndex) => ({ row, rowIndex }))
  const q = query.trim().toLowerCase()
  if (!q) return all
  if (!nameHeader) return []
  return all.filter(({ row }) => (row[nameHeader] ?? '').toLowerCase().includes(q))
}

/** Every Works List error check in one call — the "scan the list and check the errors" entry point the Errors button (App.tsx) uses. */
export function findWorksListErrors(table: ExcelTable): WorksListViolation[] {
  return [
    ...findWincodeUniquenessViolations(table),
    ...findTenderIdUniquenessViolations(table),
    ...findEcvExceedsEstimateViolations(table),
    ...findEmdInconsistencyViolations(table)
  ]
}

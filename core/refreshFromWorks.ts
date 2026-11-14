import type { ExcelTable, IssuedIntimation, StoredLetter } from './types'
import { findWorksRowByName } from './scheduleA'
import { computeWorkAmounts, tenderPercentFromRow } from './worksAmounts'

const text = (row: Record<string, string>, col: string) => (row[col] ?? '').trim()

/** Re-reads one issued letter's tender details from the Works List row with the same work name; a blank cell never overwrites what the letter already has. */
export function refreshLetter(l: StoredLetter, row: Record<string, string>): StoredLetter {
  const estimate = computeWorkAmounts(row).estimate
  const tenderId = text(row, 'Tender ID')
  const values = { ...l.values }
  if (tenderId) values.tenderId = tenderId
  if (l.kind === 'deviation') {
    try {
      const meta = JSON.parse(values.meta ?? '{}')
      if (estimate > 0) meta.estimateAmountLakhs = estimate / 100000
      if (text(row, 'Name of the Agency')) meta.agencyName = text(row, 'Name of the Agency')
      const pct = parseFloat(tenderPercentFromRow(row))
      if (Number.isFinite(pct)) meta.tenderPercent = pct
      values.meta = JSON.stringify(meta)
    } catch {
      // unreadable meta — leave it as it was
    }
  }
  return {
    ...l,
    nitNo: text(row, 'Tender Notice No') || l.nitNo,
    nitDate: text(row, 'Tender notice Date') || l.nitDate,
    agency: text(row, 'Name of the Agency') || l.agency,
    estimateRupees: estimate > 0 ? Math.round(estimate) : l.estimateRupees,
    values
  }
}

export function refreshIntimation(i: IssuedIntimation, row: Record<string, string>): IssuedIntimation {
  return { ...i, nitNo: text(row, 'Tender Notice No') || i.nitNo, agency: text(row, 'Name of the Agency') || i.agency }
}

/** Refreshes every non-deleted item whose work name is in the Works List; returns the new list and how many were found. */
export function refreshAgainstWorks<T extends { workName: string; deletedAt?: string }>(
  items: T[],
  table: ExcelTable | null,
  refresh: (item: T, row: Record<string, string>) => T,
  only: (item: T) => boolean = () => true
): { items: T[]; found: number } {
  let found = 0
  if (!table) return { items, found }
  const next = items.map((it) => {
    if (it.deletedAt || !only(it)) return it
    const hit = findWorksRowByName(table, it.workName)
    if (!hit) return it
    found++
    return refresh(it, hit.row)
  })
  return { items: next, found }
}

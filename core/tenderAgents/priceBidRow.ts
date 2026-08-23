import { PRICE_ROW, isNameContinuation, prevNonEmptyLine, toNumber } from './shared'

export interface PriceBidRow {
  ecvRupees?: number
  agencyName?: string
  tenderPercentage?: number
  contractRupees?: number
}

/**
 * Finds the price-bid table's L-1 (rank 1 / winning) row and reads every
 * value it carries in one pass — shared by the ECV, Tender Percentage,
 * Contract Value and L1 Agency agents (ecv.ts, tenderPercentage.ts,
 * contractValue.ts, l1Agency.ts), which all read from this exact row.
 *
 * Kept as one shared scan rather than each of those four agents re-running
 * its own copy of the row-matching + wrapped-name-reassembly regex: four
 * near-identical copies drifting out of sync is exactly how a fix to one
 * (e.g. rejecting the "( INR) INR)" header remnant from a wrapped name) can
 * silently fail to apply to the others. This file is the one place that
 * logic lives; each agent below just asks it for its own field.
 */
export function findL1BidRow(lines: string[]): PriceBidRow | undefined {
  // The numeric core is "<ECV> <Less|Excess> <pct> <amount> L-<rank>". The
  // bidder's company name is usually on the same line just before it, but a
  // long name wraps onto its own line(s), with the number cells landing on a
  // line *between* the two name lines (the numbers sit vertically centred
  // against the wrapped name cell) — e.g. "Kummary Renuka Devi Civil" /
  // "3133583.00 Less 11.99 2757866.40 L-1" / "Contractor". So when the number
  // row carries no company prefix, take the name from the line above, plus
  // the line below when that's a name continuation.
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim()
    const m = PRICE_ROW.exec(line)
    if (!m || Number(m[5]) !== 1) continue

    let company = line.slice(0, m.index).trim()
    if (!company) {
      const prev = prevNonEmptyLine(lines, i)
      const next = lines[i + 1]?.trim() ?? ''
      company = isNameContinuation(next) ? `${prev} ${next}`.trim() : prev
    }

    const pct = toNumber(m[3])
    return {
      ecvRupees: toNumber(m[1]),
      agencyName: company ? company.replace(/\s+/g, ' ').trim() : undefined,
      tenderPercentage: pct != null ? (/excess/i.test(m[2]) ? -pct : pct) : undefined,
      contractRupees: toNumber(m[4])
    }
  }
  return undefined
}

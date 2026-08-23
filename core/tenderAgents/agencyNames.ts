import type { NoteBidder } from '../noteSubmitted'
import { PRICE_ROW, isNameContinuation } from './shared'

/**
 * Every bidder row of the price-bid table (not just L-1), in the page's
 * order — the full data the Note Submitted bidder-comparison table needs.
 * Uses the same row shape as priceBidRow.ts's L-1-only scan
 * ("<Company> <ECV> <Less|Excess> <pct> <amount> L-<rank>") but keeps every
 * rank. "Less" shows as "(-)pct" and "Excess" as a plain "pct", matching how
 * the office notes print the quoted percentage. Returns [] when the page
 * carries no price table (e.g. the Responsiveness screen).
 */
export function detectAllBidders(lines: string[]): NoteBidder[] {
  // Anchor on the number cells, NOT on the name being on the same line: a long
  // agency name wraps, so its tail (or all of it) sits on a neighbouring line.
  // e.g. "BOBBA RAVI CHANDRA CIVIL" / "CONTRACTOR 3531887.00 Less 19.8 … L-2".
  // Reassembling the name the same way priceBidRow.ts's L-1 scan does keeps
  // every bidder, instead of dropping the wrapped one or showing only its
  // tail ("CONTRACTOR").
  const priceCore = PRICE_ROW
  // Lines already claimed as part of a bidder's (wrapped) name, so the next
  // bidder can't re-grab a previous bidder's tail — e.g. "CONTRACTOR" sitting
  // directly above "NARENDRA NAIK RAMAVATHU …" belongs to the L-2 name, not L-3.
  const consumed = new Set<number>()
  const prevUnconsumed = (i: number): number => {
    for (let j = i - 1; j >= 0; j--) {
      if (consumed.has(j)) continue
      if (lines[j].trim()) return j
    }
    return -1
  }
  const out: NoteBidder[] = []
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim()
    const m = priceCore.exec(line)
    if (!m) continue
    const prefix = line.slice(0, m.index).trim()
    let name = prefix
    if (prefix) {
      // Name tail is on this row; its head may be the line immediately above,
      // when that's a bare name fragment (not the header, not a consumed tail).
      const pj = i - 1
      if (pj >= 0 && !consumed.has(pj) && isNameContinuation(lines[pj])) {
        name = `${lines[pj].trim()} ${prefix}`
        consumed.add(pj)
      }
    } else {
      // Number-only row: the whole name sits on the neighbouring lines — the
      // nearest unclaimed line above (its head) plus the line below when that's
      // a name continuation (its tail). Both are then claimed.
      const hj = prevUnconsumed(i)
      const head = hj >= 0 ? lines[hj].trim() : ''
      if (hj >= 0) consumed.add(hj)
      const nj = i + 1
      let tail = ''
      if (nj < lines.length && !consumed.has(nj) && isNameContinuation(lines[nj])) {
        tail = lines[nj].trim()
        consumed.add(nj)
      }
      name = `${head} ${tail}`.trim()
    }
    const isExcess = /excess/i.test(m[2])
    out.push({
      sno: `${out.length + 1}.`,
      name: name.replace(/\s+/g, ' ').trim(),
      ecv: m[1].replace(/,/g, ''),
      pct: isExcess ? m[3] : `(-)${m[3]}`,
      tcv: m[4].replace(/,/g, ''),
      rank: `L-${m[5]}`
    })
  }
  return out
}

/**
 * Agent: Name of the Agencies — every bidder's company name, in the price-bid
 * table's own order (L-1, L-2, L-3, …). A thin projection of detectAllBidders
 * above, which every other bidder-table field (ECV, percentage, contract
 * value, rank) also needs — kept as one shared scan rather than a second,
 * name-only implementation that could quietly drift out of sync with it.
 */
export function detectAgencyNames(lines: string[]): string[] {
  return detectAllBidders(lines).map((b) => b.name)
}

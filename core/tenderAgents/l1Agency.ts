import { findL1BidRow } from './priceBidRow'

/**
 * Agent: L1 Agency — the L-1 (lowest / selected) bidder's company name, read
 * from the price-bid table's L-1 row (reassembled across lines when a long
 * name wraps). Undefined when the page carries no price table (e.g. the
 * Responsiveness screen, before any bid has been opened).
 */
export function detectL1Agency(lines: string[]): string | undefined {
  return findL1BidRow(lines)?.agencyName
}

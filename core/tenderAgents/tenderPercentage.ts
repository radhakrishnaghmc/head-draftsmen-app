import { findL1BidRow } from './priceBidRow'

/**
 * Agent: Tender Percentage — the L-1 bidder's quoted percentage, signed:
 * positive = below estimate ("Less"), negative = above ("Excess"). Read from
 * the price-bid table's L-1 row; undefined when the page carries no price
 * table (e.g. the Responsiveness screen, before any bid has been opened).
 */
export function detectTenderPercentage(lines: string[]): number | undefined {
  return findL1BidRow(lines)?.tenderPercentage
}

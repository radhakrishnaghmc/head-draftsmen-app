import { findL1BidRow } from './priceBidRow'

/**
 * Agent: Contract Value — the L-1 bidder's quoted amount (the awarded
 * contract value), in rupees. Read from the price-bid table's L-1 row;
 * undefined when the page carries no price table (e.g. the Responsiveness
 * screen, before any bid has been opened).
 */
export function detectContractValue(lines: string[]): number | undefined {
  return findL1BidRow(lines)?.contractRupees
}

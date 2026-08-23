import { findL1BidRow } from './priceBidRow'
import { joinLines, toNumber } from './shared'

/**
 * Agent: Estimated Contract Value, in rupees (the L1 sheet reports rupees,
 * not Lakhs) — read from the price-bid table's L-1 row when there is one,
 * else from the "Estimated Contract Value" figure on the header (present
 * even on the Responsiveness page, which has no price table at all).
 *
 * That header line's own wording varies by office: "OPEN - NCB 1593493.00"
 * (number right after the tender type, no ECV label on this line at all),
 * "OPEN 3603477.00" (no "- NCB" either), "OPEN - NCB Estimated Contract
 * Value 155560.00" (the label wedged between the tender type and the
 * number, real Nizampet Circle-58 sheets), and a Serilingampally/Ameenpur
 * office sheet ("ee-ptcu-ghmc") whose scrambled reconstruction puts the
 * number right after "Estimated Contract" with NEITHER "Value" NOR "OPEN"
 * anywhere near it: "Estimated Contract 2986947.00 OPEN - NCB …". So:
 * accept either "OPEN" (with or without "- NCB") immediately before the
 * number, OR "Estimated Contract" with "Value" only optionally following it
 * before the number.
 */
export function detectEcv(lines: string[]): number | undefined {
  const fromRow = findL1BidRow(lines)?.ecvRupees
  if (fromRow != null) return fromRow
  const joined = joinLines(lines)
  const ecv = /(?:OPEN(?:\s*-\s*NCB)?\s+|Estimated\s+Contract(?:\s+Value)?\s*\D{0,20}?)([\d,]+\.\d{2})/i.exec(joined)
  return ecv ? toNumber(ecv[1]) : undefined
}

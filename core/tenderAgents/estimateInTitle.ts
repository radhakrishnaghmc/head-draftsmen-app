import { toNumber } from './shared'

/**
 * Some offices' work titles carry a rough estimate right in the Name of
 * Work text, for at-a-glance browsing — "…(Est amt:2.00 Lakhs,Completion:1
 * Month)" or "…(Rs.5.00 lakhs)". Not universal (a scan of 305 real Nizampet
 * Circle-58 L1 pages found it on only ~4% of them — small/misc works mostly,
 * not the larger CC-road works), but when present it's a second, independent
 * source for the estimate — read straight off the title text, not the ECV
 * field — worth cross-checking the ECV agent's own answer against (see
 * crossCheck.ts) rather than trusting either source blindly.
 */
export function extractEstimateFromTitle(nameOfWork: string): number | undefined {
  const m = /(?:Est\.?\s*amt\.?|Rs\.?)\s*:?\s*([\d,]+(?:\.\d+)?)\s*(?:Lakhs?|L)\b/i.exec(nameOfWork)
  if (!m) return undefined
  const lakhs = toNumber(m[1])
  return lakhs != null ? Math.round(lakhs * 100000) : undefined
}

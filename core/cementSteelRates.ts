export interface CementSteelRate {
  slNo: number
  description: string
  datePosted: string
  token: string
  /** File extension read out of the token's own path claim (e.g. "pdf"), for a sensible default save name. */
  ext: string
}

export const CEMENT_STEEL_MONTHS = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december'
]

/**
 * The rate period a circular actually covers, as a single comparable number
 * (year * 12 + monthIndex). The department batch-uploads several months at
 * once (e.g. Dec 2025, May/Apr/Mar 2026 all "Posted" the same day), so the
 * "File Posted" date doesn't reflect the circular's own period — read the
 * month/year straight out of its description instead, and only fall back to
 * "File Posted" for the rare entry (e.g. an old quarterly one) with no plain
 * month name in its description.
 */
export function cementSteelRatePeriodKey(r: CementSteelRate): number {
  const monthMatch = /(january|february|march|april|may|june|july|august|september|october|november|december)\D{0,10}(\d{4})/i.exec(
    r.description
  )
  if (monthMatch) return Number(monthMatch[2]) * 12 + CEMENT_STEEL_MONTHS.indexOf(monthMatch[1].toLowerCase())
  const dm = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(r.datePosted.trim())
  if (dm) return Number(dm[3]) * 12 + (Number(dm[2]) - 1)
  return 0
}

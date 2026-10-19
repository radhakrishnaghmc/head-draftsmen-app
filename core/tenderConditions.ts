import type { TenderConditionsEntry } from './types'

/** Saved conditions are dropped this long after the tender's evaluation sheet is generated. */
export const TENDER_CONDITIONS_TTL_MS = 7 * 24 * 60 * 60 * 1000

/**
 * Reads the synced map tolerantly: an older build stored a bare string[] per
 * tender ID, which has no expiry yet (it gets one when a sheet is generated).
 */
export function normalizeTenderConditions(raw: unknown): Record<string, TenderConditionsEntry> {
  const out: Record<string, TenderConditionsEntry> = {}
  if (!raw || typeof raw !== 'object') return out
  for (const [id, v] of Object.entries(raw as Record<string, unknown>)) {
    if (Array.isArray(v)) out[id] = { conditions: v.map(String) }
    else if (v && typeof v === 'object' && Array.isArray((v as TenderConditionsEntry).conditions)) {
      const e = v as TenderConditionsEntry
      out[id] = typeof e.expiresAt === 'number' ? { conditions: e.conditions, expiresAt: e.expiresAt } : { conditions: e.conditions }
    }
  }
  return out
}

/** Entries whose expiry has passed are removed; returns the same object when nothing expired. */
export function purgeExpiredTenderConditions(
  map: Record<string, TenderConditionsEntry>,
  now: number
): Record<string, TenderConditionsEntry> {
  const expired = Object.keys(map).filter((id) => map[id].expiresAt != null && map[id].expiresAt! <= now)
  if (expired.length === 0) return map
  const next = { ...map }
  for (const id of expired) delete next[id]
  return next
}

/** Starts (or restarts) the one-week countdown for tender IDs whose evaluation sheet was just generated. */
export function markEvaluationGenerated(
  map: Record<string, TenderConditionsEntry>,
  tenderIds: string[],
  now: number
): Record<string, TenderConditionsEntry> {
  const next = { ...map }
  for (const id of tenderIds) {
    if (next[id]) next[id] = { ...next[id], expiresAt: now + TENDER_CONDITIONS_TTL_MS }
  }
  return next
}

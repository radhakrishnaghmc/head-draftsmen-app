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
      out[id] = { conditions: e.conditions }
      if (typeof e.expiresAt === 'number') out[id].expiresAt = e.expiresAt
      if (typeof e.nitNo === 'string') out[id].nitNo = e.nitNo
      if (typeof e.workName === 'string') out[id].workName = e.workName
      if (typeof e.ecv === 'number') out[id].ecv = e.ecv
      if (typeof e.officeKey === 'string') out[id].officeKey = e.officeKey
      if (typeof e.sharedBy === 'string') out[id].sharedBy = e.sharedBy
      if (typeof e.savedAt === 'number') out[id].savedAt = e.savedAt
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

/**
 * Whether a saved tender belongs to the current office. A tender is stamped with the
 * office it was saved from; older entries aren't, so a zone-level (SE) office is told
 * from a circle-level (EE) one by the NIT number ("/SE/" is the Superintending
 * Engineer's). Entries with neither clue stay visible everywhere.
 */
export function tenderConditionsInOffice(entry: TenderConditionsEntry, currentOfficeKey: string, zoneLevel: boolean): boolean {
  if (entry.officeKey != null) return entry.officeKey === currentOfficeKey
  if (!entry.nitNo) return true
  return /\/SE\//i.test(entry.nitNo) === zoneLevel
}

/** Drops undefined fields (Firestore rejects them) and keeps only well-formed entries, ready to write to the cloud. */
export function cleanTenderConditionsForCloud(map: Record<string, TenderConditionsEntry> | undefined): Record<string, TenderConditionsEntry> {
  const out: Record<string, TenderConditionsEntry> = {}
  for (const [id, e] of Object.entries(map ?? {})) {
    if (!e || !Array.isArray(e.conditions)) continue
    const clean: TenderConditionsEntry = { conditions: e.conditions.map(String) }
    if (typeof e.expiresAt === 'number') clean.expiresAt = e.expiresAt
    if (typeof e.nitNo === 'string') clean.nitNo = e.nitNo
    if (typeof e.workName === 'string') clean.workName = e.workName
    if (typeof e.ecv === 'number') clean.ecv = e.ecv
    if (typeof e.officeKey === 'string') clean.officeKey = e.officeKey
    if (typeof e.sharedBy === 'string') clean.sharedBy = e.sharedBy
    if (typeof e.savedAt === 'number') clean.savedAt = e.savedAt
    out[id] = clean
  }
  return out
}

/** How long a just-saved local entry is protected from a cloud update that doesn't contain it yet. */
export const FRESH_SAVE_GRACE_MS = 2 * 60 * 1000

/**
 * Applies the cloud's saved tenders to this device. The cloud is authoritative, so a tender another device
 * deleted disappears here too; a tender saved on this device moments ago (not pushed yet) is kept.
 */
export function mergeRemoteTenderConditions(
  local: Record<string, TenderConditionsEntry>,
  remote: Record<string, TenderConditionsEntry>,
  now: number
): Record<string, TenderConditionsEntry> {
  const out = { ...remote }
  for (const [id, e] of Object.entries(local)) {
    if (!(id in out) && e.savedAt != null && now - e.savedAt < FRESH_SAVE_GRACE_MS) out[id] = e
  }
  return out
}

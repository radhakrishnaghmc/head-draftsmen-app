// Pure session-slot bookkeeping for the "max concurrent logins per user"
// rule — kept free of any Firestore/network code so it can be unit tested
// directly. electron/firebaseSync.ts wraps these around a live document.

/** How many devices may be signed in to one account at the same time. */
export const MAX_CONCURRENT_SESSIONS = 5

export interface SessionSlot {
  sessionId: string
  loginAt: number
  lastSeenAt: number
  deviceLabel?: string
}

/** A slot with no heartbeat in this long is treated as dead (crashed/closed
 * without a clean logout) and no longer counts toward the limit. */
const STALE_MS = 360_000

export function liveSlots(slots: SessionSlot[], now: number): SessionSlot[] {
  return slots.filter((s) => now - s.lastSeenAt < STALE_MS)
}

export function canClaimSlot(slots: SessionSlot[], now: number): boolean {
  return liveSlots(slots, now).length < MAX_CONCURRENT_SESSIONS
}

export function claimSlot(
  slots: SessionSlot[],
  sessionId: string,
  now: number,
  deviceLabel?: string
): SessionSlot[] {
  return [...liveSlots(slots, now), { sessionId, loginAt: now, lastSeenAt: now, deviceLabel }]
}

export function releaseSlot(slots: SessionSlot[], sessionId: string): SessionSlot[] {
  return slots.filter((s) => s.sessionId !== sessionId)
}

export function touchSlot(slots: SessionSlot[], sessionId: string, now: number): SessionSlot[] {
  return slots.map((s) => (s.sessionId === sessionId ? { ...s, lastSeenAt: now } : s))
}

/**
 * Force-clears every slot for a login — the "log out other devices" recovery
 * used from the login screen once all MAX_CONCURRENT_SESSIONS are taken. A
 * stale slot already frees itself after STALE_MS with no heartbeat, but a
 * genuinely still-open session on another device won't — this lets someone
 * who already knows the password reclaim a slot immediately instead of
 * waiting, at the cost of signing every other device out.
 */
export function clearSlots(): SessionSlot[] {
  return []
}

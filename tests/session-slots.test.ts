import { describe, it, expect } from 'vitest'
import { canClaimSlot, claimSlot, releaseSlot, touchSlot, liveSlots, MAX_CONCURRENT_SESSIONS, type SessionSlot } from '../core/sessionSlots'

const now = 1_000_000

describe('sessionSlots', () => {
  const liveN = (n: number): SessionSlot[] =>
    Array.from({ length: n }, (_, i) => ({ sessionId: `s${i}`, loginAt: now, lastSeenAt: now }))

  it('allows claiming while fewer than the max slots are live', () => {
    expect(canClaimSlot([], now)).toBe(true)
    expect(canClaimSlot(liveN(MAX_CONCURRENT_SESSIONS - 1), now)).toBe(true)
  })

  it('blocks a claim once the max slots are live', () => {
    expect(canClaimSlot(liveN(MAX_CONCURRENT_SESSIONS), now)).toBe(false)
  })

  it('treats a slot with no heartbeat in 90s as dead, freeing a slot', () => {
    const stale: SessionSlot[] = [
      { sessionId: 'a', loginAt: now - 200_000, lastSeenAt: now - 100_000 },
      { sessionId: 'b', loginAt: now, lastSeenAt: now }
    ]
    expect(liveSlots(stale, now)).toEqual([{ sessionId: 'b', loginAt: now, lastSeenAt: now }])
    expect(canClaimSlot(stale, now)).toBe(true)
  })

  it('claimSlot appends the new slot and drops stale ones', () => {
    const stale: SessionSlot[] = [{ sessionId: 'a', loginAt: now - 200_000, lastSeenAt: now - 100_000 }]
    const next = claimSlot(stale, 'c', now, 'my-laptop')
    expect(next).toEqual([{ sessionId: 'c', loginAt: now, lastSeenAt: now, deviceLabel: 'my-laptop' }])
  })

  it('releaseSlot removes only the matching session', () => {
    const two: SessionSlot[] = [
      { sessionId: 'a', loginAt: now, lastSeenAt: now },
      { sessionId: 'b', loginAt: now, lastSeenAt: now }
    ]
    expect(releaseSlot(two, 'a')).toEqual([{ sessionId: 'b', loginAt: now, lastSeenAt: now }])
  })

  it('touchSlot updates only the matching session lastSeenAt', () => {
    const two: SessionSlot[] = [
      { sessionId: 'a', loginAt: now, lastSeenAt: now },
      { sessionId: 'b', loginAt: now, lastSeenAt: now }
    ]
    const later = now + 30_000
    expect(touchSlot(two, 'a', later)).toEqual([
      { sessionId: 'a', loginAt: now, lastSeenAt: later },
      { sessionId: 'b', loginAt: now, lastSeenAt: now }
    ])
  })
})

import { describe, it, expect } from 'vitest'
import { canClaimSlot, claimSlot, releaseSlot, touchSlot, liveSlots, type SessionSlot } from '../core/sessionSlots'

const now = 1_000_000

describe('sessionSlots', () => {
  it('allows claiming when there are 0 or 1 live slots', () => {
    expect(canClaimSlot([], now)).toBe(true)
    const one: SessionSlot[] = [{ sessionId: 'a', loginAt: now, lastSeenAt: now }]
    expect(canClaimSlot(one, now)).toBe(true)
  })

  it('blocks a 3rd claim when 2 slots are live', () => {
    const two: SessionSlot[] = [
      { sessionId: 'a', loginAt: now, lastSeenAt: now },
      { sessionId: 'b', loginAt: now, lastSeenAt: now }
    ]
    expect(canClaimSlot(two, now)).toBe(false)
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

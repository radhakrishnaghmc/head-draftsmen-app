import { describe, it, expect } from 'vitest'
import { makeOfflineAuthRecord, matchesOfflineAuthRecord } from '../core/offlineAuthCore'

describe('offline login credential record', () => {
  it('matches the exact password it was made from', () => {
    const record = makeOfflineAuthRecord('correct horse battery staple')
    expect(matchesOfflineAuthRecord('correct horse battery staple', record)).toBe(true)
  })

  it('rejects a wrong password', () => {
    const record = makeOfflineAuthRecord('correct horse battery staple')
    expect(matchesOfflineAuthRecord('wrong password', record)).toBe(false)
  })

  it('rejects a password differing only in case or trailing whitespace', () => {
    const record = makeOfflineAuthRecord('MyPassword123')
    expect(matchesOfflineAuthRecord('mypassword123', record)).toBe(false)
    expect(matchesOfflineAuthRecord('MyPassword123 ', record)).toBe(false)
  })

  it('never stores the plaintext password in the record', () => {
    const record = makeOfflineAuthRecord('correct horse battery staple')
    expect(record.salt).not.toContain('correct')
    expect(record.hash).not.toContain('correct')
  })

  it('uses a different salt (and therefore a different hash) each time, even for the same password', () => {
    const a = makeOfflineAuthRecord('same password')
    const b = makeOfflineAuthRecord('same password')
    expect(a.salt).not.toBe(b.salt)
    expect(a.hash).not.toBe(b.hash)
    // Both still verify correctly despite the different salts.
    expect(matchesOfflineAuthRecord('same password', a)).toBe(true)
    expect(matchesOfflineAuthRecord('same password', b)).toBe(true)
  })
})

import { describe, it, expect } from 'vitest'
import { dmyToIso } from '../core/calendar'

describe('dmyToIso', () => {
  it('converts dd.mm.yyyy to yyyy-mm-dd', () => {
    expect(dmyToIso('20.03.2026')).toBe('2026-03-20')
    expect(dmyToIso('5.7.2026')).toBe('2026-07-05')
  })

  it('returns undefined for blank or unparsable input', () => {
    expect(dmyToIso('')).toBeUndefined()
    expect(dmyToIso('not a date')).toBeUndefined()
    expect(dmyToIso('2026-03-20')).toBeUndefined()
  })
})

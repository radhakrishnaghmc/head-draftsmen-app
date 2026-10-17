import { describe, it, expect } from 'vitest'
import { agencyKey, canonicalAgencyNames } from '../core/agencyName'

describe('agencyKey', () => {
  it('merges case, M/S prefix, punctuation and spacing variants', () => {
    const k = agencyKey('AHR CONSTRUCTIONS')
    for (const v of ['M/S AHR Constructions', 'M/s. AHR Constructions', 'm/s  A.H.R. Constructions', 'Messrs AHR Constructions', ' ahr  constructions ']) {
      expect(agencyKey(v)).toBe(k)
    }
  })
  it('keeps different agencies apart, and does not eat an "ms" inside a name', () => {
    expect(agencyKey('AHR Constructions')).not.toBe(agencyKey('ECKO Design Studio'))
    expect(agencyKey('MSR Infra')).toBe('msrinfra')
    expect(agencyKey('')).toBe('')
    expect(agencyKey(undefined)).toBe('')
  })
})

describe('canonicalAgencyNames', () => {
  it('picks the most frequent variant, first seen on a tie', () => {
    const m = canonicalAgencyNames(['M/S AHR Constructions', 'AHR CONSTRUCTIONS', 'AHR CONSTRUCTIONS', '', undefined, 'ECKO Design Studio'])
    expect(m.get(agencyKey('ahr constructions'))).toBe('AHR CONSTRUCTIONS')
    expect(m.get(agencyKey('ECKO Design Studio'))).toBe('ECKO Design Studio')
    expect(m.size).toBe(2)
    expect(canonicalAgencyNames(['M/S X Co', 'X Co']).get('xco')).toBe('M/S X Co')
  })
})

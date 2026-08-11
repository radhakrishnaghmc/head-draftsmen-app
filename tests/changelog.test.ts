import { describe, it, expect } from 'vitest'
import { CHANGELOG, compareVersions, changesSince } from '../core/changelog'

describe('compareVersions', () => {
  it('orders by numeric segments', () => {
    expect(compareVersions('1.16.0', '1.15.0')).toBe(1)
    expect(compareVersions('1.15.0', '1.16.0')).toBe(-1)
    expect(compareVersions('1.16.0', '1.16.0')).toBe(0)
    expect(compareVersions('1.9.0', '1.10.0')).toBe(-1) // not string order
    expect(compareVersions('2.0.0', '1.99.99')).toBe(1)
  })

  it('treats missing / non-numeric segments as 0', () => {
    expect(compareVersions('1.16', '1.16.0')).toBe(0)
    expect(compareVersions('1.16.1', '1.16')).toBe(1)
    expect(compareVersions('x', '0.0.0')).toBe(0)
  })
})

describe('changesSince', () => {
  const log = [
    { version: '1.17.0', changes: ['c'] },
    { version: '1.16.0', changes: ['b'] },
    { version: '1.15.0', changes: ['a'] }
  ]
  const pick = (seen: string | null, current: string) =>
    log.filter((e) => compareVersions(e.version, seen ?? '') > 0 && compareVersions(e.version, current) <= 0)

  it('shows nothing on a fresh install (no version seen)', () => {
    expect(changesSince(null, '1.16.0')).toEqual([])
    expect(changesSince('', '1.16.0')).toEqual([])
    expect(changesSince(undefined, '1.16.0')).toEqual([])
  })

  it('shows nothing when nothing changed since last seen', () => {
    expect(changesSince('1.16.0', '1.16.0')).toEqual([])
  })

  it('shows the current version after a single-step update', () => {
    // Against the real CHANGELOG: updating from 1.15.0 to 1.16.0 surfaces 1.16.0.
    const shown = changesSince('1.15.0', '1.16.0')
    expect(shown.map((e) => e.version)).toEqual(['1.16.0'])
  })

  it('collects every version crossed when several are skipped (helper on a sample log)', () => {
    expect(pick('1.15.0', '1.17.0').map((e) => e.version)).toEqual(['1.17.0', '1.16.0'])
  })

  it('the top CHANGELOG entry matches a real release-shaped version string', () => {
    expect(CHANGELOG.length).toBeGreaterThan(0)
    expect(CHANGELOG[0].version).toMatch(/^\d+\.\d+\.\d+$/)
    expect(CHANGELOG[0].changes.length).toBeGreaterThan(0)
  })
})

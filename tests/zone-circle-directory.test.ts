import { describe, expect, it } from 'vitest'
import { CMC_ZONE_CIRCLES, resolveFromDirectory } from '../src/zoneCircleDirectory'

describe('resolveFromDirectory', () => {
  it('resolves a single circle mentioned in the text to its zone and CNO', () => {
    expect(resolveFromDirectory('Formation of BT road at Gajularamaram village')).toEqual({
      zone: 'Quthbullapur',
      circle: 'Gajularamaram',
      cno: '57'
    })
  })

  it('matches a multi-word circle name', () => {
    expect(resolveFromDirectory('Drain work in Allwyn Colony')).toMatchObject({
      zone: 'Kukatpally',
      circle: 'Allwyn Colony'
    })
  })

  it('is case-insensitive', () => {
    expect(resolveFromDirectory('road at MIYAPUR')).toMatchObject({ zone: 'Serilingampally', circle: 'Miyapur' })
  })

  it('returns the shared zone but no single circle when two circles of one zone are named', () => {
    const r = resolveFromDirectory('Road connecting Miyapur and Narsingi')
    expect(r.zone).toBe('Serilingampally')
    expect(r.circle).toBeUndefined()
  })

  it('resolves nothing when two circles from different zones are named', () => {
    expect(resolveFromDirectory('Road from Miyapur to Nizampet')).toEqual({})
  })

  // The explicit "<name> circle, <name> zone" designation must win over road
  // endpoints — "Miyapur" here is only the start of the road, not the circle.
  it('prefers the tagged "<name> circle, <name> zone" over incidental road endpoints', () => {
    expect(
      resolveFromDirectory(
        'laying of miyapur to bachupally road in bachupally ward of nizampet circle, quthbullapur zone cmc'
      )
    ).toEqual({ zone: 'Quthbullapur', circle: 'Nizampet', cno: '58' })
  })

  it('reads the tagged circle (zone via directory) past other locality names', () => {
    expect(resolveFromDirectory('laying of cc road in kphb to jntu in kukatpally circle, cmc')).toEqual({
      zone: 'Kukatpally',
      circle: 'Kukatpally',
      cno: '52'
    })
  })

  it('fills the zone from a "<name> zone" tag even when no circle is tagged', () => {
    expect(resolveFromDirectory('drain work in quthbullapur zone')).toEqual({
      zone: 'Quthbullapur',
      circle: undefined,
      cno: undefined
    })
  })

  it('matches only whole words, not substrings of longer place names', () => {
    // "Chintal" must not match inside "Chintalkunta".
    expect(resolveFromDirectory('Work at Chintalkunta')).toEqual({})
  })

  it('returns an empty match for text naming no known circle', () => {
    expect(resolveFromDirectory('Generic road work')).toEqual({})
    expect(resolveFromDirectory('')).toEqual({})
  })

  it('every directory entry has a distinct circle name and a CNO', () => {
    const circles = CMC_ZONE_CIRCLES.map((e) => e.circle.toLowerCase())
    expect(new Set(circles).size).toBe(circles.length)
    for (const e of CMC_ZONE_CIRCLES) expect(e.cno).toMatch(/^\d+$/)
  })
})

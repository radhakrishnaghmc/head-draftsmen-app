import { describe, it, expect } from 'vitest'
import { buildRateIndexForMatching, findMatches } from '../core/textMatch'
import type { RateEntry } from '../core/rateDatabase'

const BOILERPLATE =
  'including cost and conveyance of all materials etc complete for finished item of work as per technical specifications'

function entry(description: string, rate: string): RateEntry {
  return { description, rate, sheet: 'test', breakdown: [] }
}

describe('textMatch', () => {
  // A realistically-sized corpus matters here: with only a handful of
  // documents, IDF can't tell "shared by 3 of 3 docs" apart from "shared by
  // 3 of 300" — every entry below (bar the RCC-adjacent brick one) shares
  // the same boilerplate suffix, the way the real ~770-entry rate index does.
  const distinctItems = [
    ['Earthwork in excavation for structures', '308.31'],
    ['Brick masonry work in cement mortar', '6720.10'],
    ['Supply and fixing of aluminium windows with glass panels', '1500.00'],
    ['Plastering with cement mortar twenty mm thick', '322.78'],
    ['Filling in foundation trenches with excavated material', '1536.55'],
    ['Providing and laying stone soling for foundation', '850.00'],
    ['Dismantling of existing brick masonry walls', '410.00'],
    ['Supply and stacking of coarse sand at site', '2967.25'],
    ['White washing with two coats over one coat primer', '45.00'],
    ['Painting with synthetic enamel paint two coats', '120.00']
  ] as const
  const entries: RateEntry[] = distinctItems.map(([desc, rate]) => entry(`${desc} ${BOILERPLATE}`, rate))
  const index = buildRateIndexForMatching(entries)

  it('scores an exact match at 1.0', () => {
    const matches = findMatches(entries[0].description, index)
    expect(matches[0].entry.rate).toBe('308.31')
    expect(matches[0].score).toBeCloseTo(1, 5)
  })

  it('scores a paraphrased description highest among the correct candidate, well below 1.0', () => {
    const paraphrase = `Earth work in excavation for structure ${BOILERPLATE}`
    const matches = findMatches(paraphrase, index)
    expect(matches[0].entry.rate).toBe('308.31')
    expect(matches[0].score).toBeGreaterThan(0.5)
    expect(matches[0].score).toBeLessThan(1)
  })

  it('returns nothing for a query with no vocabulary overlap at all', () => {
    expect(findMatches('Zzyx qwibblonium fluxnetic gravitons', index)).toEqual([])
  })

  it('returns nothing for an empty query', () => {
    expect(findMatches('', index)).toEqual([])
  })
})

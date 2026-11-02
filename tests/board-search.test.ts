import { describe, it, expect } from 'vitest'
import { matchesBoardSearch } from '../core/boardSearch'

describe('matchesBoardSearch', () => {
  const work = 'Laying CC road at Yadamma Nagar in ward no. 193, Alwal Circle'
  it('matches everything on a blank query', () => {
    expect(matchesBoardSearch('', work, 'Ganta Narasimha Rao')).toBe(true)
    expect(matchesBoardSearch('   ', undefined, undefined)).toBe(true)
  })
  it('matches the work name or the agency name, case-insensitively and in any word order', () => {
    expect(matchesBoardSearch('yadamma', work, 'X')).toBe(true)
    expect(matchesBoardSearch('narasimha ganta', work, 'Ganta Narasimha Rao')).toBe(true)
    expect(matchesBoardSearch('road ganta', work, 'Ganta Narasimha Rao')).toBe(true)
  })
  it('requires every word to be present', () => {
    expect(matchesBoardSearch('yadamma kandlakoya', work, 'Ganta')).toBe(false)
  })
})

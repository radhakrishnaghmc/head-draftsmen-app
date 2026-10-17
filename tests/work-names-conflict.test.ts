import { describe, it, expect } from 'vitest'
import { workNamesConflict } from '../core/embeddingMatch'

describe('workNamesConflict', () => {
  it('flags different ward/circle numbers on otherwise identical wording', () => {
    expect(
      workNamesConflict(
        'Providing Arrangements for Bathukamma Festival (2026) at Balanagar in Ward no. 259 Moosapet Circle-53,KPZ,CMC.',
        'Making Arrangements for Bathukamma Festival at various locations in ward no -274 Bachupally in Nizampet Circle-58,Quthbullapur Zone CMC'
      )
    ).toBe(true)
  })
  it('does not flag the same ward/circle with reworded text, or missing numbers', () => {
    expect(workNamesConflict('Laying CC road in Ward No. 259, Circle-53', 'CC road laying at ward no 259 Circle 53 KPZ')).toBe(false)
    expect(workNamesConflict('Laying CC road', 'Laying CC road in Ward 12')).toBe(false)
  })
})

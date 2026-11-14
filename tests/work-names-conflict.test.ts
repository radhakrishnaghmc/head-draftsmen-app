import { describe, it, expect } from 'vitest'
import { workNamesConflict, stripWorkBoilerplate } from '../core/embeddingMatch'

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
  it('flags different house numbers in the same ward and circle', () => {
    expect(
      workNamesConflict(
        'Laying of CC Road From H.No 8-8-10/5 to 8-10/2 in Ward no-294, Dundigal Circle-59, Quthbullapur Zone, CMC.',
        'Laying of CC Road Fom H.No 6-41/1 to H.No 6-104/12/D and at H.No 6-18/2 in Ward no-294, Dundigal Circle-59, Quthubullapur Zone, CMC (Reserved for SC only)'
      )
    ).toBe(true)
    expect(workNamesConflict('CC road at H.No 6-41/1 Ward 294', 'Laying CC road from H.No. 6-41/1 to 6-50 ward no 294')).toBe(false)
  })
})

describe('shared-template names', () => {
  const t = (place: string) => `Laying of CC Road ${place} in Ward no-294, Dundigal Circle-59, Quthubullapur Zone, CMC`
  it('compares only the place-specific part when the template is identical', () => {
    expect(workNamesConflict(t('at Laxmi Kirana Colony Bahadurpally'), t('at Jahanger Basthi Internal Roads'))).toBe(true)
    expect(workNamesConflict(t('at Jahanger Basthi Internal Roads'), t('at Jahangir Bsthi internal road'))).toBe(false)
  })
  it('does not apply when the wording differs', () => {
    expect(workNamesConflict(t('at Jahanger Basthi'), 'Construction of SW Drain at Jahanger Basthi in Ward no-294, Circle-59, Zone, CMC')).toBe(false)
  })
})

describe('stripWorkBoilerplate', () => {
  it('keeps only the place-specific part of a work name', () => {
    const out = stripWorkBoilerplate(
      'Laying of CC Road Fom H.No 6-41/1 to H.No 6-104/12/D and At H.No 6-18/2 and at Internal Roads Jahanger Bsthi in Ward no-294,Bahadurpally in Dundigal Circle-59,Quthubullapur Zone,CMC (Reserved for SC only)'
    )
    expect(out).not.toMatch(/laying|cc road|ward|circle|zone|cmc|reserved/i)
    expect(out).toMatch(/6-41\/1/)
    expect(out).toMatch(/Jahanger/)
  })
  it('falls back to the full text when nothing is left', () => {
    expect(stripWorkBoilerplate('Construction of CC Road')).toBe('Construction of CC Road')
  })
})

import { describe, expect, it } from 'vitest'
import { normalizeTenderConditions, tenderConditionsInOffice } from '../core/tenderConditions'

const SE = 'CMC|Quthbullapur||'.replace(/\|\|$/, '|')
const EE = 'CMC|Quthbullapur|Nizampet'

describe('tenderConditionsInOffice', () => {
  it('uses the stamped office key when present', () => {
    const e = { conditions: ['a'], officeKey: EE }
    expect(tenderConditionsInOffice(e, EE, false)).toBe(true)
    expect(tenderConditionsInOffice(e, SE, true)).toBe(false)
  })
  it('infers an unstamped entry from the NIT number', () => {
    const se = { conditions: ['a'], nitNo: 'NIT No. 28/SE/QBZ/CMC/2026-27' }
    const ee = { conditions: ['a'], nitNo: 'NIT No. 29/2/DB/EE/Nizampet Circle-58/CMC/2026-27' }
    expect(tenderConditionsInOffice(se, SE, true)).toBe(true)
    expect(tenderConditionsInOffice(se, EE, false)).toBe(false)
    expect(tenderConditionsInOffice(ee, EE, false)).toBe(true)
    expect(tenderConditionsInOffice(ee, SE, true)).toBe(false)
  })
  it('keeps entries with no clue visible everywhere', () => {
    const e = { conditions: ['a'] }
    expect(tenderConditionsInOffice(e, EE, false)).toBe(true)
    expect(tenderConditionsInOffice(e, SE, true)).toBe(true)
  })
  it('normalize keeps the office key', () => {
    expect(normalizeTenderConditions({ '1': { conditions: ['a'], officeKey: EE } })['1'].officeKey).toBe(EE)
  })
})

import { cleanTenderConditionsForCloud, mergeRemoteTenderConditions } from '../core/tenderConditions'

describe('tender conditions cloud sync', () => {
  it('cleans entries for Firestore: no undefined fields', () => {
    const out = cleanTenderConditionsForCloud({ '1': { conditions: ['a'], nitNo: undefined, workName: 'w', ecv: undefined, officeKey: EE, savedAt: 5 } })
    expect(Object.keys(out['1']).sort()).toEqual(['conditions', 'officeKey', 'savedAt', 'workName'])
    expect(Object.values(out['1']).includes(undefined as never)).toBe(false)
  })
  it('a remote delete removes the tender here', () => {
    const local = { '1': { conditions: ['a'], savedAt: 1000 }, '2': { conditions: ['b'] } }
    const merged = mergeRemoteTenderConditions(local, { '2': { conditions: ['b'] } }, 10 * 60 * 1000)
    expect(Object.keys(merged)).toEqual(['2'])
  })
  it('keeps a tender saved here moments ago that the cloud does not have yet', () => {
    const now = 1_000_000
    const local = { '1': { conditions: ['a'], savedAt: now - 5000 } }
    expect(Object.keys(mergeRemoteTenderConditions(local, {}, now))).toEqual(['1'])
  })
  it('remote wins when both have the tender', () => {
    const merged = mergeRemoteTenderConditions({ '1': { conditions: ['old'], savedAt: 9 } }, { '1': { conditions: ['new'] } }, 10)
    expect(merged['1'].conditions).toEqual(['new'])
  })
})

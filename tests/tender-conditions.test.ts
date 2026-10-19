import { describe, expect, it } from 'vitest'
import {
  TENDER_CONDITIONS_TTL_MS,
  markEvaluationGenerated,
  normalizeTenderConditions,
  purgeExpiredTenderConditions
} from '../core/tenderConditions'

const NOW = 1_800_000_000_000

describe('tender conditions expiry', () => {
  it('keeps conditions with no expiry until a sheet is generated', () => {
    const map = { '741286': { conditions: ['a', 'b'] } }
    expect(purgeExpiredTenderConditions(map, NOW + 10 * TENDER_CONDITIONS_TTL_MS)).toBe(map)
  })

  it('starts a one-week countdown only for the tender whose sheet was generated', () => {
    const map = { '1': { conditions: ['a'] }, '2': { conditions: ['b'] } }
    const next = markEvaluationGenerated(map, ['1'], NOW)
    expect(next['1'].expiresAt).toBe(NOW + TENDER_CONDITIONS_TTL_MS)
    expect(next['2'].expiresAt).toBeUndefined()
  })

  it('drops an entry once its week is up, and leaves the others', () => {
    const map = markEvaluationGenerated({ '1': { conditions: ['a'] }, '2': { conditions: ['b'] } }, ['1'], NOW)
    expect(Object.keys(purgeExpiredTenderConditions(map, NOW + TENDER_CONDITIONS_TTL_MS - 1))).toEqual(['1', '2'])
    expect(Object.keys(purgeExpiredTenderConditions(map, NOW + TENDER_CONDITIONS_TTL_MS))).toEqual(['2'])
  })

  it('ignores an id with nothing saved', () => {
    expect(markEvaluationGenerated({}, ['9'], NOW)).toEqual({})
  })

  it('reads the older bare-array shape', () => {
    expect(normalizeTenderConditions({ '1': ['a'], '2': { conditions: ['b'], expiresAt: 5 }, '3': 'x' })).toEqual({
      '1': { conditions: ['a'] },
      '2': { conditions: ['b'], expiresAt: 5 }
    })
  })
})

import { describe, expect, it } from 'vitest'
import { verifyAmountMath } from '../core/documentVerify'

describe('EMD check', () => {
  const base = { ecv: 1536433, contract: 1536433, tenderPct: 0 }
  it('accepts the EMD printed with paise', () => {
    const emd = (1536433 * 0.015).toFixed(2)
    const [w, p] = emd.split('.')
    const text = `ECV Rs.15,36,433 EMD Rs.${Number(w).toLocaleString('en-IN')}.${p}`
    expect(verifyAmountMath({ ...base, docText: text, emdPct: 0.015 }).some((i) => i.field === 'EMD')).toBe(false)
  })
  it('still flags a missing EMD', () => {
    expect(verifyAmountMath({ ...base, docText: 'ECV Rs.15,36,433', emdPct: 0.015 }).some((i) => i.field === 'EMD')).toBe(true)
  })
})

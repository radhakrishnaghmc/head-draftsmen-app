import { describe, it, expect } from 'vitest'
import { resolveIntimationValue } from '../core/intimationFill'
import type { IntimationNotice } from '../core/intimationNotice'
import type { TenderEvaluation } from '../core/tenderEvaluationPdf'

// Minimal notice/pdf — only the amount + name fields the EMD/ASD branch reads.
const notice = {} as IntimationNotice
const pdf = (over: Partial<TenderEvaluation>) => ({ ...over }) as TenderEvaluation
const RESERVED = 'Construction of RCC box drain (Reserved for SC only)'
const OPEN = 'Providing CC road at Balaji Nagar'

describe('Intimation EMD/ASD placeholder', () => {
  it('keeps ASD for a RESERVED work quoted >25% below (EMD exempted, ASD still charged)', () => {
    // ECV 31,20,000 · 32% below → ASD = (32−25)% × ECV = 2,18,400
    const row = { 'Name of the work': RESERVED, ECV: '3120000' }
    const out = resolveIntimationValue('EMD 1.5%', notice, pdf({ tenderPercentage: 32 }), row)
    expect(out).toBe('Rs. 1 ½ Rs.Exempted,ASD Rs.218400/-')
  })

  it('shows only "Exempted" for a RESERVED work at 25% or less (no ASD due)', () => {
    const row = { 'Name of the work': RESERVED, ECV: '3120000' }
    const out = resolveIntimationValue('EMD 1.5%', notice, pdf({ tenderPercentage: 10 }), row)
    expect(out).toBe('Rs. 1 ½ Rs.Exempted/-')
  })

  it('still appends ASD for an OPEN work quoted >25% below', () => {
    const row = { 'Name of the work': OPEN, ECV: '3120000' }
    const out = resolveIntimationValue('EMD 1.5%', notice, pdf({ tenderPercentage: 32 }), row)
    // EMD @1.5% = floor(3120000*0.015) = 46800
    expect(out).toBe('Rs. 1 ½ Rs.46800,ASD Rs.218400/-')
  })

  it('the standalone ASD placeholder also carries ASD for a reserved work', () => {
    const row = { 'Name of the work': RESERVED, ECV: '3120000' }
    expect(resolveIntimationValue('ASD', notice, pdf({ tenderPercentage: 32 }), row)).toBe('ASD Rs.218400/-')
  })
})

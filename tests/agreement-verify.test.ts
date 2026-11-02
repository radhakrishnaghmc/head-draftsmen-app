import { describe, it, expect } from 'vitest'
import { verifyDocFigures, verifyWorkFigures, type DocFigures } from '../core/agreementVerify'

const f: DocFigures = {
  workName: 'Construction of CC road at Ward 5',
  estimateLakhs: 12.5,
  ecv: 1000000,
  contract: 800000,
  tenderPct: 20,
  wincode: 'WC123',
  completionMonths: '3',
  reserved: false
}
const labels = ['Name of the work', 'Estimate Amount', 'ECV', 'Contract Amount', 'Tender Percentage', 'Wincode', 'Period Completion']
const goodText =
  'Construction of CC road at Ward 5 Rs. 12.50 Lakhs ECV Rs.10,00,000.00 Rs.8,00,000.00 (-) 20.00%-Less WC123 3 Months'

describe('verifyDocFigures', () => {
  it('passes when every figure is present', () => {
    expect(verifyDocFigures(goodText, labels, f)).toEqual([])
  })
  it('flags a wrong estimate, wincode and name', () => {
    const bad = goodText.replace('12.50', '99.00').replace('WC123', 'X').replace('CC road', 'BT road')
    const fields = verifyDocFigures(bad, labels, f).map((i) => i.field)
    expect(fields).toEqual(expect.arrayContaining(['Estimate', 'Wincode', 'Name of Work']))
  })
  it('ignores figures the template has no placeholder for', () => {
    expect(verifyDocFigures('nothing here', ['Name of the work'], { ...f, workName: '' }).length).toBe(1)
  })
  it('checks EMD 1%, 1.5% and ASD', () => {
    const l = ['EMD 1 percent', 'EMD 1.5 percent', 'ASD Details']
    const ok = verifyDocFigures('10,000 15,000 1,00,000', l, { ...f, tenderPct: 35 })
    expect(ok).toEqual([])
    expect(verifyDocFigures('nil', l, { ...f, tenderPct: 35 }).map((i) => i.field)).toEqual(['EMD 1%', 'EMD 1.5%', 'ASD'])
    expect(verifyDocFigures('Exempted', l, { ...f, reserved: true, tenderPct: 10 })).toEqual([])
  })
})

describe('verifyWorkFigures', () => {
  const base = { ...f, fieldsEstimateLakhs: 12.5, worksRowFound: true, monitoringWorks: [] }
  it('flags documents filled with a different estimate than the database', () => {
    const r = verifyWorkFigures({ ...base, fieldsEstimateLakhs: 40 })
    expect(r.issues.some((i) => i.field === 'Estimate')).toBe(true)
  })
  it('flags ECV >= estimate and contract above ECV on a Less tender', () => {
    const r = verifyWorkFigures({ ...base, ecv: 2000000, contract: 2100000 })
    expect(r.issues.map((i) => i.field)).toEqual(expect.arrayContaining(['ECV', 'Contract Amount']))
  })
  it('flags blank wincode / period and reports skipped Monitoring Format', () => {
    const r = verifyWorkFigures({ ...base, wincode: '', completionMonths: '' })
    expect(r.issues.map((i) => i.field)).toEqual(expect.arrayContaining(['Wincode', 'Period of Completion']))
    expect(r.skipped.some((s) => s.startsWith('Monitoring Format'))).toBe(true)
  })
  it('flags a Works List EMD that disagrees with 1% of the L-1 ECV', () => {
    const r = verifyWorkFigures({ ...base, rowEmdAsd: { emd1: '9,000' } })
    expect(r.issues.some((i) => i.field === 'EMD 1%')).toBe(true)
  })
  it('is clean for consistent figures', () => {
    expect(verifyWorkFigures({ ...base, monitoringWorks: undefined }).issues).toEqual([])
  })
})

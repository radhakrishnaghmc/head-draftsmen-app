import { describe, expect, it } from 'vitest'
import { parseTenderEvaluation } from '../core/tenderEvaluationPdf'

// pdf.js reconstructs these lines from the portal's Commercial Evaluation
// page — the header's two-column layout interleaves "Tender ID …" into the
// middle of the NIT No, exactly as the real page does.
const COMMERCIAL_LINES = [
  'Commercial Evaluation',
  'NIT No. 12/DB/EE/Nizampet Circle-',
  'Tender ID 717574',
  '58/CMC/2026-27 item 1 Dated:15.07.2026',
  'Notice Number',
  'Name of Work Junction Improvement in Aleap Circle in Ward no 276 Pragathi Nagar in Nizampet circle-58, Quthbullapur Zone, CMC',
  'Tender Category Works Tender Evaluation Type Percentage',
  'Estimated Contract',
  'OPEN - NCB 1593493.00',
  'Company Name Estimated Contract Value ( INR) Excess/Less Percentage(%) Amount ( INR) Rank Select',
  'M V S CONSTRUCTIONS 1593493.00 Less 11.11 1416455.93 L-1',
  'T VASANTHA REDDY 1593493.00 Less 1.26 1573414.99 L-2'
]

const RESPONSIVENESS_LINES = [
  'Preliminary Responsiveness',
  'NIT No. 12/DB/EE/Nizampet Circle-',
  'Tender ID 717574',
  '58/CMC/2026-27 item 1 Dated:15.07.2026',
  'Notice Number',
  'Name of Work Junction Improvement in Aleap Circle in Ward no 276 Pragathi Nagar, CMC',
  'Tender Category Works Tender Evaluation Type Percentage',
  'OPEN - NCB 1593493.00',
  'SlNo. Company Name EMD Details Transaction Fee',
  '1 T VASANTHA REDDY',
  '2 M V S CONSTRUCTIONS'
]

describe('parseTenderEvaluation', () => {
  it('extracts every field from the commercial-evaluation (L1) page', () => {
    const r = parseTenderEvaluation(COMMERCIAL_LINES)
    expect(r.tenderId).toBe('717574')
    expect(r.noticeNo).toBe('12/DB/EE/Nizampet Circle-58/CMC/2026-27')
    expect(r.nameOfWork).toBe(
      'Junction Improvement in Aleap Circle in Ward no 276 Pragathi Nagar in Nizampet circle-58, Quthbullapur Zone, CMC'
    )
    expect(r.ecvRupees).toBe(1593493)
    expect(r.l1AgencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.tenderPercentage).toBeCloseTo(11.11, 2)
    expect(r.contractRupees).toBeCloseTo(1416455.93, 2)
  })

  it('takes the L-1 row, not L-2, for the winning bid', () => {
    const r = parseTenderEvaluation(COMMERCIAL_LINES)
    expect(r.l1AgencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.contractRupees).not.toBeCloseTo(1573414.99, 2) // that's L-2
  })

  it('signs an "Excess" quote negative so contract-amount math stays correct', () => {
    const r = parseTenderEvaluation([
      'Company Name Estimated Contract Value ( INR) Excess/Less Percentage(%) Amount ( INR) Rank Select',
      'SOME BUILDERS 1000000.00 Excess 5.00 1050000.00 L-1'
    ])
    expect(r.tenderPercentage).toBe(-5)
    expect(r.contractRupees).toBe(1050000)
  })

  it('still gets Tender ID / NIT / ECV / Name of Work from the responsiveness page (no price table)', () => {
    const r = parseTenderEvaluation(RESPONSIVENESS_LINES)
    expect(r.tenderId).toBe('717574')
    expect(r.noticeNo).toBe('12/DB/EE/Nizampet Circle-58/CMC/2026-27')
    expect(r.ecvRupees).toBe(1593493)
    expect(r.nameOfWork).toContain('Junction Improvement in Aleap Circle')
    expect(r.l1AgencyName).toBeUndefined()
    expect(r.tenderPercentage).toBeUndefined()
    expect(r.contractRupees).toBeUndefined()
  })
})

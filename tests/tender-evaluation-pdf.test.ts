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
    expect(r.noticeDate).toBe('15.07.2026')
    expect(r.nameOfWork).toBe(
      'Junction Improvement in Aleap Circle in Ward no 276 Pragathi Nagar in Nizampet circle-58, Quthbullapur Zone, CMC'
    )
    expect(r.ecvRupees).toBe(1593493)
    expect(r.l1AgencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.tenderPercentage).toBeCloseTo(11.11, 2)
    expect(r.contractRupees).toBeCloseTo(1416455.93, 2)
  })

  // A long Name of Work wraps around its own label: value part 1, then the
  // "Name of Work" label alone, then value part 2, then the next field. Taken
  // from a real GHMC L1 form (Stage Selected Form). The old single-line regex
  // captured only "Circle-58, … Works Percentage" — a fragment that then
  // mis-matched a completely different work.
  const WRAPPED_WORKNAME_LINES = [
    'Commercial Evaluation',
    'E1/06/23/DB/EE/Nizampet Circle-',
    'Tender ID 710682',
    '58/CMC/2026-27, dt: 18.06.2026',
    'Notice Number',
    'Laying of CC Road from Tirumala Enclave to Rasari Apartment in Ward 7 in NMC (Ward No 276, Pragathi nagar Nizampet',
    'Name of Work',
    'Circle-58, Quthbullapur Zone CMC) (Reserved for SC)',
    'Works Percentage',
    'Estimated Contract',
    'Tender Type OPEN - NCB 2242283.00',
    'Company Name Excess/Less Percentage(%) Rank Select',
    'CHALLAPURAM SREE',
    '2242283.00 Less 10.96 1996528.78 L-1',
    'DEVAYANI'
  ]

  // Real Gajularamaram L1 (Stage Selected Form): the Tender ID's value lands a
  // row ABOVE its label ("699549" then "Tender ID"), sitting between the NIT's
  // two wrapped halves, and the NIT tail reads "Dt.…" not "Dated". The old
  // parser mis-read Tender ID as "57" and left "699549" inside the NIT.
  const GAJULARAMARAM_LINES = [
    'Commercial Evaluation',
    'Enquiry/IFB/Tender Notice NIT No.02/DB/EE/Gajularamaram Circle-',
    '699549',
    'Tender ID',
    '57/QBZ/CMC/2026-27 Dt.07.05.2026 (Item No.02)',
    'Number',
    'Maintenance of Circle office by supplying of water bubbles for a period of (12) months and basic office equipments, Sanitary Items in',
    'Name of Work',
    'Gajularamaram circle office of Gajularamaram Circle-57, CMC.',
    'Tender Category Works Tender Evaluation Type Percentage',
    'Company Name Estimated Contract Value ( INR) Excess/Less Percentage(%) Amount ( INR) Rank Select',
    'Sri Duggi Parvathalu 105528.00 Less 21.34 83008.32 L-1'
  ]

  it('reads Tender ID and NIT when the Tender ID value lands above its label, between the NIT halves', () => {
    const r = parseTenderEvaluation(GAJULARAMARAM_LINES)
    expect(r.tenderId).toBe('699549')
    expect(r.noticeNo).toBe('02/DB/EE/Gajularamaram Circle-57/QBZ/CMC/2026-27')
    expect(r.noticeDate).toBe('07.05.2026')
    expect(r.l1AgencyName).toBe('Sri Duggi Parvathalu')
  })

  it('drops an "Enquiry/IFB/Tender <id>" label + stray Tender ID wedged into the NIT', () => {
    // Real layout where pdf.js interleaves the right column's label and value
    // between the NIT's wrapped halves.
    const r = parseTenderEvaluation([
      'Commercial Evaluation',
      'Enquiry/IFB/Tender Notice NIT No.02/DB/EE/Gajularamaram Circle- Enquiry/IFB/Tender 699588 57/QBZ/CMC/2026-27 (Item No.01)',
      'Name of Work Toilets',
      'M V S CONSTRUCTIONS 1000000.00 Less 5.00 950000.00 L-1'
    ])
    expect(r.noticeNo).toBe('02/DB/EE/Gajularamaram Circle-57/QBZ/CMC/2026-27')
  })

  it('trims a lowercase "date:" / work-name tail the value regex over-captured', () => {
    const r = parseTenderEvaluation([
      'Commercial Evaluation',
      'NIT No.13/DB/EE/Gajularamaram Circle-57/CMC/2026-27 date:4.07.2026 Arrangement of 51 Computer Operators on hire'
    ])
    expect(r.noticeNo).toBe('13/DB/EE/Gajularamaram Circle-57/CMC/2026-27')
  })

  it('reads a Name of Work whose value wraps around its label, not just the tail fragment', () => {
    const r = parseTenderEvaluation(WRAPPED_WORKNAME_LINES)
    expect(r.nameOfWork).toBe(
      'Laying of CC Road from Tirumala Enclave to Rasari Apartment in Ward 7 in NMC (Ward No 276, Pragathi nagar Nizampet Circle-58, Quthbullapur Zone CMC) (Reserved for SC)'
    )
    expect(r.nameOfWork).not.toContain('Works Percentage')
    expect(r.l1AgencyName).toBe('CHALLAPURAM SREE DEVAYANI')
  })

  it('reads a Name of Work whose value wraps onto TWO lines above the label (title + tail)', () => {
    // Real SVS Infra L1: the title spans two lines before the "Name of Work"
    // label, then "(Reserved for ST)" after it. The old parser kept only the
    // second line ("…under Municipal General Funds…"), dropping the title, so
    // the Work Order matched a completely different work.
    const r = parseTenderEvaluation([
      'Commercial Evaluation',
      'Enquiry/IFB/Tender E1/06/17/DB/EE/Nizampet Circle-58/CMC/2026-27',
      'Notice Number',
      'Laying of CC Road From SVR Infra SV Sadan to Plot no 59,60 to Sri Sai Datta Residency in Nizampet Municipal Corporation',
      'under Municipal General Funds 2025-26 (ward No 274, Bachupally Nizampet Circle-58, Quthbullapur Zone CMC)',
      'Name of Work',
      '(Reserved for ST)',
      'Works Percentage',
      'M V S CONSTRUCTIONS 1195243.00 Less 5.00 1135480.85 L-1'
    ])
    expect(r.nameOfWork).toBe(
      'Laying of CC Road From SVR Infra SV Sadan to Plot no 59,60 to Sri Sai Datta Residency in Nizampet Municipal Corporation under Municipal General Funds 2025-26 (ward No 274, Bachupally Nizampet Circle-58, Quthbullapur Zone CMC) (Reserved for ST)'
    )
  })

  it('takes the L-1 row, not L-2, for the winning bid', () => {
    const r = parseTenderEvaluation(COMMERCIAL_LINES)
    expect(r.l1AgencyName).toBe('M V S CONSTRUCTIONS')
    expect(r.contractRupees).not.toBeCloseTo(1573414.99, 2) // that's L-2
  })

  // A long L-1 company name wraps onto its own line(s) in the price-bid cell,
  // and pdf.js lands the number row on the line *between* the two name lines
  // (the numbers sit vertically centred against the wrapped name cell). Taken
  // from a real Gajularamaram L1 form.
  const WRAPPED_NAME_LINES = [
    'Price Bid Details /Commercial Stage',
    'Estimated Contract Value ( Amount (',
    'Company Name Excess/Less Percentage(%) Rank Select',
    'INR) INR)',
    'Kummary Renuka Devi Civil',
    '3133583.00 Less 11.99 2757866.40 L-1',
    'Contractor',
    'Back Save & Continue Reject Tender'
  ]

  it('reassembles an L-1 company name that wraps around the number row', () => {
    const r = parseTenderEvaluation(WRAPPED_NAME_LINES)
    expect(r.l1AgencyName).toBe('Kummary Renuka Devi Civil Contractor')
    expect(r.ecvRupees).toBe(3133583)
    expect(r.tenderPercentage).toBeCloseTo(11.99, 2)
    expect(r.contractRupees).toBeCloseTo(2757866.4, 2)
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
    expect(r.noticeDate).toBe('15.07.2026')
    expect(r.ecvRupees).toBe(1593493)
    expect(r.nameOfWork).toContain('Junction Improvement in Aleap Circle')
    expect(r.l1AgencyName).toBeUndefined()
    expect(r.tenderPercentage).toBeUndefined()
    expect(r.contractRupees).toBeUndefined()
  })
})

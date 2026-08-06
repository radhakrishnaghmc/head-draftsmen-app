import { describe, it, expect } from 'vitest'
import { summarizeNonResponsiveness, buildNoteSubmittedHtml, type NoteSubmittedData } from '../core/noteSubmitted'

// The exact lines pdf.js reconstructs from the portal's "List of Bidders Made
// Non-Responsive" sheet (Tender 717652, 2 rejected bidders). Includes the page
// chrome that also carries the "Non-Responsive" keyword — the title, the
// section header and the "Please Select … Non-Responsiveness" instruction —
// which used to be miscounted as bidders (the sheet reported "(5) rejected").
const NON_RESPONSIVE_2_BIDDERS = [
  'Welcome to ee-grrc-ghmc Profile | Training Manuals | Logout',
  'List of Bidders Made Non-Responsive',
  'Current Tender Details',
  'Name of Work',
  'List of Bidders Made Non-Responsive /Commercial Stage Please Enter Comments',
  'Company Name Registration Digital Scanned Criminal Others Comments *',
  'EMD',
  'SRI TULJA',
  'Non Responsive due to',
  'BHAVANI',
  'low Bid Capacity',
  'CONSTRUCTIONS',
  'Non Responsive Low',
  'SVS INFRA',
  'bid Capacity and',
  'Please Select Only such reasons for Non-Responsiveness,strictly based on NIT/Tender conditions * Indicates Mandatory',
  'Please Select the reason for the Disqualification'
]

describe('summarizeNonResponsiveness', () => {
  it('counts only the rejected bidders, not the page title / header / instruction', () => {
    const { count } = summarizeNonResponsiveness(NON_RESPONSIVE_2_BIDDERS)
    expect(count).toBe(2)
  })

  it('leaves the reason blank for several bidders (their reasons do not concatenate readably)', () => {
    const { detail } = summarizeNonResponsiveness(NON_RESPONSIVE_2_BIDDERS)
    expect(detail).toBe('')
  })

  it('extracts a clean single reason when exactly one bidder is rejected', () => {
    const oneBidder = NON_RESPONSIVE_2_BIDDERS.filter(
      (l) => !/SVS INFRA|Non Responsive Low|bid Capacity and/.test(l)
    )
    const { count, detail } = summarizeNonResponsiveness(oneBidder)
    expect(count).toBe(1)
    expect(detail).toBe('low Bid Capacity')
  })

  it('returns zero when the sheet flags no rejected bidders', () => {
    expect(summarizeNonResponsiveness(['Current Tender Details', 'Company Name'])).toEqual({
      count: 0,
      detail: ''
    })
  })
})

function noteData(over: Partial<NoteSubmittedData>): NoteSubmittedData {
  return {
    body: 'CMC',
    circle: 'Nizampet',
    workName: 'Laying of UGD',
    estimateLakhs: '35.99',
    asDate: '',
    financialYear: '2026-27',
    tenderNoticeNo: 'NIT/1',
    tenderNoticeDate: '14.07.2026',
    nitNo: 'NIT/1',
    nitDate: '14.07.2026',
    newspapers: 'Andhra Jyothi',
    qualificationNote: '',
    rejectedCount: 0,
    bidders: [
      { sno: '1.', name: 'L SURENDER', ecv: '3599124.00', pct: '(-)26.99', tcv: '2627720.43', rank: 'L-1' },
      { sno: '2.', name: 'NARENDRA NAIK', ecv: '3599124.00', pct: '(-)25.05', tcv: '2697543.44', rank: 'L-2' },
      { sno: '3.', name: 'ARUNA', ecv: '3599124.00', pct: '(-)25', tcv: '2699343.00', rank: 'L-3' },
      { sno: '4.', name: 'SRINIVASA', ecv: '3599124.00', pct: '(-)19.01', tcv: '2914930.53', rank: 'L-4' }
    ],
    l1Name: 'L SURENDER',
    l1PctText: '(-)26.99',
    l1Tcv: '2627720.43',
    intimationDate: '29.07.2026',
    reservation: 'ST',
    l1PctNumber: 26.99,
    ecvRupees: 3599124,
    receiptNo: '',
    receiptDate: '',
    ...over
  }
}

describe('buildNoteSubmittedHtml — participant / rejected counts', () => {
  it('counts participants as qualified + rejected (not just the price table)', () => {
    // 4 in the price table + 2 rejected = 6 participated.
    const html = buildNoteSubmittedHtml(noteData({ rejectedCount: 2, qualificationNote: 'low bid capacity' }))
    expect(html).toContain('(6) bidders have participated')
    expect(html).toContain('In that (2) bidders rejected due to low bid capacity and (4) qualified as follows.')
  })

  it('reads as a plain participant count when nothing was rejected', () => {
    const html = buildNoteSubmittedHtml(noteData({ rejectedCount: 0 }))
    expect(html).toContain('(4) bidders have participated as follows.')
    expect(html).not.toContain('rejected')
  })

  it('uses singular "bidder" for a single rejection', () => {
    const html = buildNoteSubmittedHtml(noteData({ rejectedCount: 1 }))
    expect(html).toContain('(5) bidders have participated')
    expect(html).toContain('In that (1) bidder rejected and (4) qualified as follows.')
  })
})

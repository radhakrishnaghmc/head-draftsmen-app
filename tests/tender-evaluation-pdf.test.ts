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

  it("reads the footer's Server Time date (bottom-right) as the server date, normalised to dd.mm.yyyy", () => {
    // The L1 sheet's bottom-right footer, split across lines by pdf.js the way
    // the real page wraps the date above the time.
    const r = parseTenderEvaluation([...COMMERCIAL_LINES, 'a, India. All Server Time: 02/07/2026', '03:59:31 PM'])
    expect(r.serverDate).toBe('02.07.2026')
  })

  it('leaves the server date undefined when no Server Time footer is present', () => {
    expect(parseTenderEvaluation(COMMERCIAL_LINES).serverDate).toBeUndefined()
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

  it('drops the same wedged Tender ID from a "C-<n>" short circle code, not just "Circle-<n>"', () => {
    // Real EE-III/DB layout ("13/DB/EE-III/CNL/C-54/QBZ/CMC/2026-27") — a
    // division office identifies its circle as "C-54", not a named-place
    // "Circle-54", but pdf.js still wedges the Tender ID between the wrapped
    // halves the same way. Reported bug: with only the "Circle-" anchor, the
    // wedge survived uncleaned and made this NIT compare as a different work
    // from the same NIT read off the LOA (which has no such wedge).
    const r = parseTenderEvaluation([
      'Current Tender Details',
      'NIT No.13/DB/EE-III/CNL/C-',
      'Enquiry/IFB/Tender',
      '723483 54/QBZ/CMC/2026-27, Dt:31.07.2026',
      'Tender ID',
      'Notice Number',
      '(Item no. 1)',
      'Name of Work Development of Muslim graveyard',
      'VADDI MANISH REDDY 1091236.00 Less 3.01 1058389.80 L-1'
    ])
    expect(r.noticeNo).toBe('13/DB/EE-III/CNL/C-54/QBZ/CMC/2026-27')
    expect(r.tenderId).toBeUndefined()
  })

  it('reads the SE-office Tender ID when the NIT date year sits right before the label', () => {
    // Real SE (zone) L1 layout: "… Dt. 27.07.2026 Tender ID 722264 (Item No.11)".
    // The "2026" of the NIT date abuts the label; the old \d+ grabbed it (leftmost).
    const r = parseTenderEvaluation([
      'Current Tender Details Enquiry/IFB/Tender Notice NIT No. 15/SE/QBZ/CMC/2026-27, Dt. 27.07.2026 Tender ID 722264 (Item No.11) Number',
      'Name of Work Laying of Footpath',
      'J.J. Constructions 1000000.00 Less 5.00 950000.00 L-1'
    ])
    expect(r.tenderId).toBe('722264')
  })

  it('reads the notice No & date from an "Enquiry/IFB/Tender … Notice Number" line with no "NIT No." label', () => {
    // Real NIT.08 layout: the field is labelled only "Enquiry/IFB/Tender …
    // Notice Number", never "NIT No.", and the Tender ID splits the wrapped
    // "…/2026-  27" tail. The notice number (and its Dt: date) must still come
    // through so the Note Submitted form isn't left blank.
    const r = parseTenderEvaluation([
      'Enquiry/IFB/Tender 08/DB/EE/Nizampet Circle-58/CMC/2026-',
      'Tender ID 711763',
      'Notice Number 27(Item No.01),Dt:25.06.2026',
      'Name of Work Maintenance of CC Roads'
    ])
    expect(r.noticeNo).toBe('08/DB/EE/Nizampet Circle-58/CMC/2026-27')
    expect(r.noticeDate).toBe('25.06.2026')
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
    // Reported bug: the "E1/06/" item-number prefix was dropped, truncating
    // the notice number to "23/DB/EE/…" instead of the full "E1/06/23/DB/EE/…".
    expect(r.noticeNo).toBe('E1/06/23/DB/EE/Nizampet Circle-58/CMC/2026-27')
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
    expect(r.noticeNo).toBe('E1/06/17/DB/EE/Nizampet Circle-58/CMC/2026-27')
  })

  it('keeps the full "E1/06/" item-number prefix on a real GHMC Nizampet Circle-58 L1 sheet', () => {
    // Real Stage Selected Form (Commercial Evaluation) — the notice number
    // wraps mid-code with no "NIT No." label at all, just "Enquiry/IFB/Tender
    // … Notice Number". Reported bug: the app showed the wrong/truncated
    // tender number in the Intimation and Agreement workspace.
    const r = parseTenderEvaluation([
      'Commercial Evaluation',
      'Current Tender Details',
      'Enquiry/IFB/Tender',
      'E1/06/11/DB/EE/Nizampet Circle-',
      'Tender ID 710618',
      '58/CMC/2026-27, dt: 18.06.2026',
      'Notice Number',
      'Laying of CC road from Sunrise Homes Plot No 172 to Komaram Bheem Park in ward No 22 in Nizampet Municipal',
      'Corporation under Municipal general funds 2025-26 (Ward No 276, Pragathi nagar Nizampet Circle-58, Quthbullapur Zone',
      'Name of Work',
      'CMC)',
      'Works Percentage',
      'Bid Submission Start Bid Submission Closing',
      '22/06/2026 07:00 PM 29/06/2026 03:30 PM',
      'Date & Time Date',
      'Company Name Estimated Contract Value ( INR) Excess/Less Percentage(%) Amount ( INR) Rank Select',
      'SP CONSTRCUTIONS 1033505.00 Less 22 806133.90 L-1'
    ])
    expect(r.noticeNo).toBe('E1/06/11/DB/EE/Nizampet Circle-58/CMC/2026-27')
    expect(r.noticeDate).toBe('18.06.2026')
    expect(r.bidClose).toBe('29/06/2026 03:30 PM')
    expect(r.contractRupees).toBe(806133.9)
  })

  it('drops a stray "Number" column-header fragment wedged between an "(Item No.N)" tag and the title', () => {
    // Real SE L1 sheet (8-L1.pdf): the page stacks "(Item No.8)" then a bare
    // "Number" line then the title, right above the "Name of Work" label —
    // the old parser walked past both and printed "(Item No.8) Number Laying
    // of CC roads…" as the work name (the "(Item No.8)" tags are cleaned up
    // downstream by stripItemNoTag, but nothing removed the leaked "Number").
    const r = parseTenderEvaluation([
      'Commercial Evaluation',
      'Current Tender Details',
      'Enquiry/IFB/Tender Notice NIT No. 15/SE/QBZ/CMC/2026-27, Dt. 27.07.2026',
      'Tender ID 722257',
      '(Item No.8)',
      'Number',
      'Laying of CC roads in internal lanes of Bathukamma banda 01/113/ER/446, S.no 1247, 1-113/A/ER/466, 1-113/ER/530/1, 1-113/A/ER/151C,',
      'Name of Work',
      'graveyard) of ward no. 277, Mahadevapuram, Gajularamaram Circle-57, CMC (Reserved for SC Only) (Item No.8)',
      'Tender Category Works Tender Evaluation Type Percentage',
      'Tender Type OPEN - NCB Estimated Contract Value 6262019.00',
      'J.J. Constructions 6262019.00 Less 25 4696514.25 L-1'
    ])
    expect(r.nameOfWork).not.toMatch(/^number/i)
    expect(r.nameOfWork).not.toContain('Number Laying')
    expect(r.nameOfWork).toBe(
      'Laying of CC roads in internal lanes of Bathukamma banda 01/113/ER/446, S.no 1247, 1-113/A/ER/466, 1-113/ER/530/1, 1-113/A/ER/151C, graveyard) of ward no. 277, Mahadevapuram, Gajularamaram Circle-57, CMC (Reserved for SC Only) (Item No.8)'
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

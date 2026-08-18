import { describe, it, expect } from 'vitest'
import {
  summarizeNonResponsiveness,
  buildNoteSubmittedHtml,
  noteSubmittedFromRow,
  tenderPctMagnitude,
  type NoteSubmittedData,
  type NonRespLine
} from '../core/noteSubmitted'

const NAME_X = 40
const COMMENT_X = 600

// Positioned lines mirroring the portal's "List of Bidders Made Non-Responsive"
// sheet: a company-name column on the far left (NAME_X) and a Comments column on
// the far right (COMMENT_X). y increases downward, as pdfToPositionedLines emits.
const header = (y: number): NonRespLine => ({
  text: 'Company Name Registration Digital Scanned Criminal Others Comments *',
  x: NAME_X,
  y
})
const name = (text: string, y: number): NonRespLine => ({ text, x: NAME_X, y })
const comment = (text: string, y: number): NonRespLine => ({ text, x: COMMENT_X, y })
const footer = (y: number): NonRespLine => ({
  text: 'Please Select Only such reasons for Non-Responsiveness,strictly based on NIT/Tender conditions',
  x: NAME_X,
  y
})

// Two bidders (SRI TULJA BHAVANI CONSTRUCTIONS, SVS INFRA) — name and comment
// lines interleave in y as the taller name cell centres against the comment.
const TWO_BIDDERS: NonRespLine[] = [
  header(100),
  name('SRI TULJA', 120),
  comment('Non Responsive due to', 122),
  name('BHAVANI', 135),
  comment('low Bid Capacity', 137),
  name('CONSTRUCTIONS', 150),
  name('SVS INFRA', 180),
  comment('Non Responsive Low', 182),
  comment('bid Capacity and', 197),
  footer(220)
]

describe('summarizeNonResponsiveness', () => {
  it('counts bidders from the table structure, not the page chrome', () => {
    expect(summarizeNonResponsiveness(TWO_BIDDERS).count).toBe(2)
  })

  it('leaves the reason blank for several bidders (they do not concatenate readably)', () => {
    expect(summarizeNonResponsiveness(TWO_BIDDERS).detail).toBe('')
  })

  it('extracts a clean single reason when exactly one bidder is rejected', () => {
    const one: NonRespLine[] = [
      header(100),
      name('SRI TULJA', 120),
      comment('Non Responsive due to', 122),
      name('BHAVANI', 135),
      comment('low Bid Capacity', 137),
      name('CONSTRUCTIONS', 150),
      footer(200)
    ]
    const { count, detail } = summarizeNonResponsiveness(one)
    expect(count).toBe(1)
    expect(detail).toBe('low Bid Capacity')
  })

  it('counts a bidder even when the evaluator misspells "responsive"', () => {
    const typo: NonRespLine[] = [
      header(100),
      name('MSR', 120),
      comment('NOT RESPOSONVIE DUE', 122),
      comment('TO LOW BID', 137),
      name('CONSTRUCTIONS', 150),
      footer(200)
    ]
    const { count, detail } = summarizeNonResponsiveness(typo)
    expect(count).toBe(1)
    expect(detail).toBe('LOW BID')
  })

  it('returns zero when there is no distinct comments column', () => {
    expect(summarizeNonResponsiveness([header(100), footer(120)])).toEqual({ count: 0, detail: '' })
  })

  // Real GRRC sheet (Tender 720492): three agencies, each with a SINGLE-line
  // comment and a two-line name. Comment gaps are all row-height (~33), so the
  // old comment-median threshold saw one long block → 1 bidder. The name column's
  // wrapped lines (gap ~13) give the true line height, splitting all three.
  const THREE_SINGLE_LINE_COMMENTS: NonRespLine[] = [
    header(0),
    name('KAILA SHIVA', 45),
    comment('not uploaded turnover', 46),
    name('KUMAR', 58),
    name('MSR', 78),
    comment('low bid capacity', 79),
    name('CONSTRUCTIONS', 91),
    name('Shri Karthikeya', 111),
    comment('low bid capacity', 112),
    name('Enterprises', 124),
    footer(157)
  ]

  it('counts three agencies each with a single-line comment (was collapsing to one)', () => {
    expect(summarizeNonResponsiveness(THREE_SINGLE_LINE_COMMENTS).count).toBe(3)
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

  it('leaves a fill-in blank after "dt:" when no admin-sanction date is set', () => {
    const html = buildNoteSubmittedHtml(noteData({ asDate: '' }))
    expect(html).toContain('dt: _____________ under general budget')
  })

  it('prints the admin-sanction date after "dt:" when set', () => {
    const html = buildNoteSubmittedHtml(noteData({ asDate: '29.05.2026' }))
    expect(html).toContain('dt: 29.05.2026 under general budget')
  })

  it('justifies every note paragraph', () => {
    const html = buildNoteSubmittedHtml(noteData({}))
    expect(html).not.toContain('<p style="margin:0 0 6px;line-height:1.45">') // no un-justified paragraphs
    expect(html.includes('text-align:justify')).toBe(true)
  })

  it('prints a full-line blank for the newspapers when none entered, not a default', () => {
    const html = buildNoteSubmittedHtml(noteData({ newspapers: '' }))
    expect(html).toContain(`published in ${'_'.repeat(65)}.`)
    expect(html).not.toContain('Andhra Jyothi')
  })

  it('prints the newspapers when entered', () => {
    const html = buildNoteSubmittedHtml(noteData({ newspapers: 'Eenadu and The Hindu' }))
    expect(html).toContain('published in Eenadu and The Hindu.')
  })

  it('gives the EMD Online Receipt No a full-line blank when unset (20-char number)', () => {
    const html = buildNoteSubmittedHtml(noteData({ receiptNo: '' }))
    expect(html).toContain(`Online Receipt No: ${'_'.repeat(45)} `)
  })

  it('shows ASD for a RESERVED work quoted >25% below — EMD exempted, ASD still charged, with receipt', () => {
    const html = buildNoteSubmittedHtml(
      noteData({ reservation: 'SC', l1PctNumber: 32, ecvRupees: 3120000, receiptNo: 'R-123', receiptDate: '05.08.2026' })
    )
    expect(html).toContain('the EMD is Exempted and submitted ASD amount of Rs.')
    expect(html).toContain('vide Online Receipt No: R-123 Dt: 05.08.2026')
  })

  it('shows only "EMD is Exempted" for a reserved work at 25% or less (no ASD due)', () => {
    const html = buildNoteSubmittedHtml(noteData({ reservation: 'SC', l1PctNumber: 20, ecvRupees: 3120000 }))
    expect(html).toContain('the EMD is Exempted')
    expect(html).not.toContain('ASD amount')
  })

  it('first-line indents every note paragraph (a tab of space)', () => {
    const html = buildNoteSubmittedHtml(noteData({}))
    // Body paragraphs open with an em-space indent (html-to-docx ignores CSS text-indent).
    expect(html).toContain('text-align:justify">&emsp;&emsp;&emsp;')
    expect(html).toContain('text-align:left">&emsp;&emsp;&emsp;')
  })
})

describe('tender percentage magnitude (drives ASD)', () => {
  it('reads the magnitude from every stored shape', () => {
    expect(tenderPctMagnitude('(-)32.00')).toBe(32)
    expect(tenderPctMagnitude('32 % Less')).toBe(32)
    expect(tenderPctMagnitude('(-) 32%-Less')).toBe(32)
    expect(tenderPctMagnitude('-32')).toBe(32)
    expect(tenderPctMagnitude('')).toBeNull()
    expect(tenderPctMagnitude(null)).toBeNull()
  })

  it('keeps l1PctNumber (and thus ASD) alive when the row percentage is formatted — was NaN→null before', () => {
    const d = noteSubmittedFromRow({
      'Name of the work': 'RCC drain (Reserved for SC only)',
      ECV: '3120000',
      'Tender Percentage': '32 % Less'
    })
    expect(d.l1PctNumber).toBe(32)
  })
})

describe('noteSubmittedFromRow prefers the uploaded L-1 / Online Intimation over a mismatched Works List row', () => {
  // Reproduces the real bug: the Works List match is name-similarity based
  // (falls back to embeddings), so it can land on an unrelated-but-similar
  // row that carries ITS OWN name/ECV/agency/NIT — a different work than the
  // one actually described by the uploaded L-1 sheet / Online Intimation.
  // Every field Note Submitted shows must come from the uploads, not this row.
  const wrongRow = {
    'Name of the work': 'Laying of CC road from Shubamkaree temple to babu Jagjivan park in Pragathi nagar ward no 276',
    Circle: 'Nizampet',
    'Amount of estimate': '35.00',
    ECV: '1450000',
    'Tender Percentage': '9.20',
    'Name of the Agency': 'SOME OTHER CONTRACTOR',
    'Tender Notice No': '7/DB/EE/Old',
    'Tender notice Date': '01.01.2026'
  }
  const pdf = {
    nameOfWork: 'Laying of CC Road from RGK STP to Children Park in Bhandari Layout ward no 275',
    ecvRupees: 2080503,
    tenderPercentage: 18.66,
    contractRupees: 1692281.14,
    l1AgencyName: 'NANDU CONSTRUCTIONS',
    noticeNo: '15/DB/EE/Nizampet-58',
    noticeDate: '12.08.2026',
    serverDate: '13.08.2026'
  }

  it('takes the work name from the L-1 sheet, not the matched row', () => {
    const d = noteSubmittedFromRow(wrongRow, pdf, {})
    expect(d.workName).toBe(pdf.nameOfWork)
    expect(d.workName).not.toContain('Shubamkaree')
  })

  it('takes ECV, tender %, contract value and agency from the L-1 sheet, not the matched row', () => {
    const d = noteSubmittedFromRow(wrongRow, pdf, {})
    expect(d.ecvRupees).toBe(pdf.ecvRupees)
    expect(d.l1PctNumber).toBe(pdf.tenderPercentage)
    expect(d.l1Tcv).toBe(pdf.contractRupees.toFixed(2))
    expect(d.l1Name).toBe(pdf.l1AgencyName)
  })

  it('takes the NIT No/date and Intimation date from the L-1 sheet, not the matched row', () => {
    const d = noteSubmittedFromRow(wrongRow, pdf, {})
    expect(d.tenderNoticeNo).toBe(pdf.noticeNo)
    expect(d.nitNo).toBe(pdf.noticeNo)
    expect(d.tenderNoticeDate).toBe(pdf.noticeDate)
    expect(d.nitDate).toBe(pdf.noticeDate)
    expect(d.intimationDate).toBe(pdf.serverDate)
  })

  it('the Online Intimation outranks even the L-1 sheet for agency/ECV/contract/NIT when both are present', () => {
    const notice = {
      agencyName: 'INTIMATION AGENCY',
      ecvRupees: 3000000,
      contractRupees: 2500000,
      nitNo: '99/DB/EE/Intimation',
      nitDate: '20.08.2026'
    }
    const d = noteSubmittedFromRow(wrongRow, pdf, notice)
    expect(d.l1Name).toBe(notice.agencyName)
    expect(d.ecvRupees).toBe(notice.ecvRupees)
    expect(d.l1Tcv).toBe(notice.contractRupees.toFixed(2))
    expect(d.tenderNoticeNo).toBe(notice.nitNo)
    expect(d.tenderNoticeDate).toBe(notice.nitDate)
  })

  it('still falls back to the row when no L-1 / Intimation was uploaded', () => {
    const d = noteSubmittedFromRow(wrongRow, {}, {})
    expect(d.workName).toContain('Shubamkaree')
    expect(d.l1Name).toBe('SOME OTHER CONTRACTOR')
    expect(d.tenderNoticeNo).toBe('7/DB/EE/Old')
  })
})

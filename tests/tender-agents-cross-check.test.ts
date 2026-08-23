import { describe, expect, it } from 'vitest'
import { runTenderAgents, sameTender } from '../core/tenderAgents/crossCheck'
import { extractEstimateFromTitle } from '../core/tenderAgents/estimateInTitle'

describe('extractEstimateFromTitle', () => {
  it('reads "(Est amt:2.00 Lakhs,...)" as rupees', () => {
    expect(extractEstimateFromTitle('…CMC(Est amt:2.00 Lakhs,Completion:1 Month)')).toBe(200000)
  })

  it('reads "(Rs.5.00 lakhs)" as rupees', () => {
    expect(extractEstimateFromTitle('Artistic painting at park (Rs.5.00 lakhs)')).toBe(500000)
  })

  it('returns undefined when the title carries no amount at all', () => {
    expect(extractEstimateFromTitle('Laying of CC Road from Nayara Petrol Bunk to Bhandari Layout')).toBeUndefined()
  })
})

describe('runTenderAgents — cross-checks ECV against a work title\'s own rough estimate', () => {
  // Real NIT.03(FY 2026-27)NZ/1).Item No.01-700210 Responsiveness page: the
  // title states "(Est amt:2.00 Lakhs…)" (₹200,000) while the page's own
  // "Estimated Contract Value" field reads 155560.00 — a real, roughly-30%
  // disagreement (the title figure is a rounded headline estimate, not the
  // precise ECV) that's worth surfacing, not silently picking one.
  const LINES_WITH_TITLE_ESTIMATE = [
    'Preliminary Responsiveness',
    'Current Tender Details',
    'Enquiry/IFB/Tender Notice NIT No. 03/DB/EE/Nizampet Circle-58/CMC/2026-',
    '700210',
    'Tender ID',
    '27(Item No.01)',
    'Number',
    'Laying of underground drainage pipe line from existing toilet in Pragathi nagar MPPS to main gate in Pragathi nagar ward',
    'Name of Work',
    '276,Quthbullapur Zone ,CMC(Est amt:2.00 Lakhs,Completion:1 Month)',
    'Tender Category Works Tender Evaluation Type Percentage',
    'Tender Type OPEN - NCB Estimated Contract Value 155560.00'
  ]

  it('keeps the page\'s own ECV as the value, but warns about the disagreement', () => {
    const r = runTenderAgents(LINES_WITH_TITLE_ESTIMATE)
    expect(r.ecvRupees).toBe(155560) // the field's own value wins, never silently overwritten
    expect(r.warnings.some((w) => w.includes('disagree'))).toBe(true)
    expect(r.warnings.some((w) => w.includes('1,55,560'))).toBe(true) // Indian digit grouping, not Western
    expect(r.warnings.some((w) => w.includes('2,00,000'))).toBe(true)
  })

  it('falls back to the title\'s estimate, with a warning, when the page has no ECV field at all', () => {
    const noEcvField = LINES_WITH_TITLE_ESTIMATE.filter((l) => !/Estimated Contract Value|OPEN/.test(l))
    const r = runTenderAgents(noEcvField)
    expect(r.ecvRupees).toBe(200000)
    expect(r.warnings.some((w) => w.includes('using the estimate mentioned in the work title'))).toBe(true)
  })

  it('adds no ECV warning at all when the title carries no estimate to compare against', () => {
    // Real NIT.17/726879 sheet — no amount ever mentioned in this title.
    const r = runTenderAgents([
      'Name of Work',
      'Laying of CC Road from Nayara Petrol Bunk to Bhandari Layout Road Junction',
      'Tender Category Works Tender Evaluation Type Percentage',
      'Tender Type OPEN 1521909.00'
    ])
    expect(r.ecvRupees).toBe(1521909)
    expect(r.warnings).toEqual([])
  })
})

describe('sameTender — Tender ID / NIT No as the shared identity key across documents', () => {
  // Two REAL, DIFFERENT tenders from the same NIT.17 folder (726878 and
  // 726879) — same circle, same NIT series, issued the same day, item
  // numbers one apart. If anything were going to be mistaken for "the same
  // tender", it's these two.
  const TENDER_726878 = runTenderAgents([
    'NIT No',
    'Enquiry/IFB/Tender',
    '726878 17/DB/EE/NizampetCircle58/CMC/2026-27',
    'Tender ID',
    'Notice Number',
    'ITEM 4 ,Dated:18.08.2026',
    'Name of Work',
    'Laying of Approach CC Road to Sub lanes in Pragathi Nagar Main Road'
  ])
  const TENDER_726879 = runTenderAgents([
    'NIT No',
    'Enquiry/IFB/Tender',
    '726879 17/DB/EE/NizampetCircle58/CMC/2026-27',
    'Tender ID',
    'Notice Number',
    'ITEM 5 Dated:18.08.2026',
    'Name of Work',
    'Laying of CC Road from Nayara Petrol Bunk to Bhandari Layout Road Junction'
  ])

  it('a document compared against itself is always the same tender', () => {
    expect(sameTender(TENDER_726878, TENDER_726878)).toBe(true)
  })

  it('two different real tenders sharing the same NIT series are correctly told apart', () => {
    expect(TENDER_726878.tenderId).not.toBe(TENDER_726879.tenderId)
    expect(sameTender(TENDER_726878, TENDER_726879)).toBe(false)
  })

  it('an Intimation letter (no Tender ID field at all) still matches its own L1 sheet via NIT No', () => {
    const intimationForSameTender = { tenderId: undefined, noticeNo: '17/DB/EE/NizampetCircle58/CMC/2026-27' }
    expect(sameTender(TENDER_726879, intimationForSameTender)).toBe(true)
  })

  it('two documents with nothing in common to compare are never assumed to match', () => {
    expect(sameTender({}, {})).toBe(false)
    expect(sameTender({ tenderId: '726879' }, {})).toBe(false)
  })
})

describe('runTenderAgents — identity-shape cross-check between Tender ID and NIT No', () => {
  it('flags a NIT No that carries none of the expected "/DB/" or "/SE/" segment', () => {
    // cleanNit's raw-fallback path (used when neither of its own two shape
    // regexes match) can pass through un-validated text — this simulates
    // that fallback firing on a garbled/unexpected NIT line.
    const r = runTenderAgents(['NIT No. some garbled text with no real code in it Dated'])
    expect(r.warnings.some((w) => w.includes('NIT No') && w.includes('/DB/'))).toBe(true)
  })

  it('adds no identity warning for a normal, well-formed real tender', () => {
    const r = runTenderAgents([
      'NIT No',
      'Enquiry/IFB/Tender',
      '726879 17/DB/EE/NizampetCircle58/CMC/2026-27',
      'Tender ID',
      'Notice Number',
      'ITEM 5 Dated:18.08.2026'
    ])
    expect(r.warnings.filter((w) => w.includes('Tender ID') || w.includes('NIT No'))).toEqual([])
  })
})

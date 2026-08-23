import { describe, expect, it } from 'vitest'
import { detectTenderId } from '../core/tenderAgents/tenderId'
import { detectNitNoAndDate } from '../core/tenderAgents/nitNoAndDate'
import { detectNameOfWork } from '../core/tenderAgents/nameOfWork'
import { detectEcv } from '../core/tenderAgents/ecv'
import { detectTenderPercentage } from '../core/tenderAgents/tenderPercentage'
import { detectContractValue } from '../core/tenderAgents/contractValue'
import { detectBidSubmissionStartDate, detectBidSubmissionEndDate } from '../core/tenderAgents/bidSubmissionDates'
import { detectAgencyNames, detectAllBidders } from '../core/tenderAgents/agencyNames'
import { detectL1Agency } from '../core/tenderAgents/l1Agency'
import { detectAgencyAddress } from '../core/tenderAgents/agencyAddress'
import { linesFromOcr, linesFromExcelRows } from '../core/tenderAgents/inputAdapters'
import { runTenderAgents } from '../core/tenderAgents/crossCheck'

// Real Nizampet Circle-58 NIT.17 L1 sheet (Tender ID 726879) — the exact
// pdf.js-reconstructed lines, 4 bidders, and a 3-line-wrapped price-bid
// header ("( INR) INR)" landing alone right above the L-1 row). Every agent
// below is exercised against this ONE real sheet to prove each one reads its
// own field correctly and independently — no agent here depends on another
// agent's output.
const L1_LINES = [
  'Welcome to ee-grrc-ghmc Profile | Training Manuals | Logout',
  'Dashboard Tender Creation Tender Evaluation LOA',
  'Commercial Evaluation',
  'Current Tender Details',
  'NIT No',
  'Enquiry/IFB/Tender',
  '726879 17/DB/EE/NizampetCircle58/CMC/2026-27',
  'Tender ID',
  'Notice Number',
  'ITEM 5 Dated:18.08.2026',
  'Laying of CC Road from Nayara Petrol Bunk to Bhandari Layout Road Junction in Bhandari Layout ward no 275 in Nizampet',
  'Name of Work',
  'Circle -58 ,Quthbullapur Zone ,CMC (1ST RECALL)',
  'Tender Category Works Tender Evaluation Type Percentage',
  'Estimated Contract',
  'OPEN 1521909.00',
  'Tender Type',
  'Value',
  'Bid Submission Start Bid Submission Closing',
  '19/08/2026 05:20 PM 22/08/2026 02:30 PM',
  'Date & Time Date',
  'Price Bid Details /Commercial Stage',
  'Estimated Contract Value Amount (',
  'Company Name Excess/Less Percentage(%) Rank Select',
  '( INR) INR)',
  'NANDU CONSTRUCTIONS 1521909.00 Less 8.66 1390111.68 L-1',
  'KOLAN RAMA KRISHNA REDDY 1521909.00 Less 7.29 1410961.83 L-2',
  'SHIVARATHRI SANDEEP 1521909.00 Less 2.99 1476403.92 L-3',
  'CHIRANJEVI ALAKUNTLA WORKS',
  '1521909.00 Less 2.10 1489948.91 L-4',
  'CONTRACTOR',
  'Back Save & Continue Reject Tender'
]

// Real Intimation / Letter of Acceptance letter for the same tender — the
// only document that ever carries the winning agency's postal address.
const INTIMATION_LINES = [
  'DATE: Saturday, August 22, 2026',
  'To',
  'NANDU CONSTRUCTIONS',
  'H No 1-2-3, Nizampet Village',
  'Medchal Malkajgiri -500090',
  'Telangana',
  'Sir/Madam,',
  'This is notify you that the bid submitted by you for execution of the NIT No',
  '17/DB/EE/NizampetCircle58/CMC/2026-27 ITEM 5 ,Dated:18.08.2026 at contract price of Rs.',
  '1390111.68 ( Thirteen Lakh Ninety Thousand One Hundred and Eleven Rupees Sixty Eight Paisa)',
  'Company Name Estimated Contract Value Corpus Fund @ 0.04 %',
  'NANDU CONSTRUCTIONS 1521909.00 608.00',
  'Yours Faithfully'
]

describe('tender agents — each reads one field off the real NIT.17 (726879) L1 sheet', () => {
  it('Agent 1 — Tender ID', () => {
    expect(detectTenderId(L1_LINES)).toBe('726879')
  })

  it('Agent 2 — NIT No + NIT Date', () => {
    const r = detectNitNoAndDate(L1_LINES)
    expect(r.noticeNo).toBe('17/DB/EE/NizampetCircle58/CMC/2026-27')
    expect(r.noticeDate).toBe('18.08.2026')
  })

  it('Agent 3 — Name of Work', () => {
    expect(detectNameOfWork(L1_LINES)).toBe(
      'Laying of CC Road from Nayara Petrol Bunk to Bhandari Layout Road Junction in Bhandari Layout ward no 275 in Nizampet Circle -58 ,Quthbullapur Zone ,CMC (1ST RECALL)'
    )
  })

  it('Agent 4 — Estimated Contract Value (ECV)', () => {
    expect(detectEcv(L1_LINES)).toBe(1521909)
  })

  it('Agent 5 — Tender Percentage', () => {
    expect(detectTenderPercentage(L1_LINES)).toBeCloseTo(8.66, 2)
  })

  it('Agent 6 — Contract Value', () => {
    expect(detectContractValue(L1_LINES)).toBeCloseTo(1390111.68, 2)
  })

  it('Agent 7 — Bid Submission Start Date', () => {
    expect(detectBidSubmissionStartDate(L1_LINES)).toBe('19/08/2026 05:20 PM')
  })

  it('Agent 8 — Bid Submission (download) End Date', () => {
    expect(detectBidSubmissionEndDate(L1_LINES)).toBe('22/08/2026 02:30 PM')
  })

  it('Agent 9 — Name of the Agencies (every bidder, in table order)', () => {
    expect(detectAgencyNames(L1_LINES)).toEqual([
      'NANDU CONSTRUCTIONS',
      'KOLAN RAMA KRISHNA REDDY',
      'SHIVARATHRI SANDEEP',
      'CHIRANJEVI ALAKUNTLA WORKS CONTRACTOR'
    ])
    // Reported bug: the "( INR) INR)" wrapped header remnant landed right
    // above the first bidder row and got glued onto its name.
    expect(detectAgencyNames(L1_LINES)[0]).not.toMatch(/inr/i)
  })

  it('detectAllBidders backs the Name of the Agencies agent with the full row data the Note Submitted table needs', () => {
    const bidders = detectAllBidders(L1_LINES)
    expect(bidders).toHaveLength(4)
    expect(bidders[0]).toMatchObject({ name: 'NANDU CONSTRUCTIONS', rank: 'L-1', pct: '(-)8.66' })
  })

  it('Agent 10 — L1 Agency', () => {
    expect(detectL1Agency(L1_LINES)).toBe('NANDU CONSTRUCTIONS')
  })

  it('Agent 11 — Address of the Agency (read from the Intimation letter, not the L1 sheet)', () => {
    // The L1 sheet itself carries no address at all — confirms this agent
    // must be run against the Intimation letter, a different document.
    expect(detectAgencyAddress(L1_LINES).address).toBeUndefined()

    const r = detectAgencyAddress(INTIMATION_LINES)
    expect(r.agencyName).toBe('NANDU CONSTRUCTIONS')
    expect(r.address).toBe('H No 1-2-3, Nizampet Village, Medchal Malkajgiri -500090, Telangana')
  })
})

describe('tender agents — independence: each still returns undefined/[] gracefully when its own field is absent, without needing any other field present', () => {
  it('every agent handles a page with none of the fields it looks for', () => {
    const empty = ['Some unrelated page text', 'with nothing recognizable on it']
    expect(detectTenderId(empty)).toBeUndefined()
    expect(detectNitNoAndDate(empty)).toEqual({})
    expect(detectNameOfWork(empty)).toBeUndefined()
    expect(detectEcv(empty)).toBeUndefined()
    expect(detectTenderPercentage(empty)).toBeUndefined()
    expect(detectContractValue(empty)).toBeUndefined()
    expect(detectBidSubmissionStartDate(empty)).toBeUndefined()
    expect(detectBidSubmissionEndDate(empty)).toBeUndefined()
    expect(detectAgencyNames(empty)).toEqual([])
    expect(detectL1Agency(empty)).toBeUndefined()
    expect(detectAgencyAddress(empty)).toEqual({})
  })
})

// A bulk scan of every real L1/Intimation PDF under a live circle's own
// Google Drive folder (NIZAMPET-58, 305 L1 sheets + 57 Intimation letters)
// surfaced three more real-world layout variants the synthetic test fixtures
// above never happened to cover — all on the Preliminary Responsiveness page
// (no price table yet) or in the NIT No's own printed code. Each block below
// is the exact reconstructed text of one real failing file.
describe('tender agents — real-world variants found by a bulk scan of NIZAMPET-58', () => {
  it('Agent 4 (ECV) — Responsiveness page reading "OPEN <amount>" with no "- NCB" at all', () => {
    // Real: NIT.06-26 WORKS/.../2).710538.../Stage Selected Form 1.pdf
    expect(detectEcv(['Tender Category Tender Evaluation Type', 'Estimated Contract', 'Tender Type OPEN 3603477.00', 'Value'])).toBe(
      3603477
    )
  })

  it('Agent 4 (ECV) — Responsiveness page with "Estimated Contract Value" wedged between "OPEN - NCB" and the number', () => {
    // Real: NIT.03(FY 2026-27)NZ/1).Item No.01-700210/Stage Selected Form 1.pdf
    expect(
      detectEcv(['Tender Type OPEN - NCB Estimated Contract Value 155560.00', 'Bid Submission Start Date Bid Submission Closing'])
    ).toBe(155560)
  })

  it('Agent 1 (Tender ID) — id sits before the NIT tail\'s "<circleNo>/[QBZ/]CMC/<year>", not a fresh "/DB/EE/"', () => {
    // Real: NIT.02(FY-026-27)/1.iTEM .01-699966/L1.pdf — the id lands right
    // before the WRAPPED TAIL of a NIT No the label-anchored branch already
    // consumed the "/DB/EE/…Circle-" head of, so no "/DB/EE/" follows it here.
    const lines = [
      'NIT No.02/DB/EE/Nizampet Circle-',
      'Enquiry/IFB/Tender',
      '699966 58/QBZ/CMC/2026-27 Dt.08.05.2026 (Item',
      'Tender ID',
      'Notice Number',
      'No.01)'
    ]
    expect(detectTenderId(lines)).toBe('699966')
  })

  it('Agent 2 (NIT No) — a real code that omits the "/CMC/" segment entirely', () => {
    // Real: NIT.06-26 WORKS/.../1).710536.../Stage Selected Form 1.pdf — this
    // office's older "Completed-26 works" NITs print no corporation segment.
    const lines = [
      'Enquiry/IFB/Tender',
      'E1/06/01/DB/EE/Nizampetcircle-58/2026-27,',
      'Tender ID 710536',
      'dt:18.06.2026',
      'Notice Number'
    ]
    expect(detectNitNoAndDate(lines).noticeNo).toBe('E1/06/01/DB/EE/Nizampetcircle-58/2026-27')
  })

  it('Agent 2 (NIT No) — a real code that omits BOTH "/EE/" and "/CMC/"', () => {
    // Real: NIT.06-26 WORKS/.../2).710538.../Stage Selected Form 1.pdf
    const lines = ['Enquiry/IFB/Tender', 'E1/06/02/DB/Nizampet Circle-58/2026-27,Dt', 'Tender ID 710538']
    expect(detectNitNoAndDate(lines).noticeNo).toBe('E1/06/02/DB/Nizampet Circle-58/2026-27')
  })
})

// Every agent only ever reads `lines: string[]` — it never knows or cares
// whether those lines came from a digital PDF's text layer, a photo run
// through this app's OCR engine, or a spreadsheet's rows. These two blocks
// prove that by running the SAME agents, unchanged, against the SAME real
// tender's data reshaped as each of those other sources.
describe('tender agents — work unchanged against an OCR\'d photo (via linesFromOcr)', () => {
  // The exact L1_LINES fixture above, but shuffled into a random order and
  // given each a `top` position — exactly the shape electron/ocr.ts's line
  // detector returns for a photographed page, where detection order has
  // nothing to do with reading order. linesFromOcr must restore reading
  // order before any agent can find anything.
  const scrambledOcrLines = [...L1_LINES]
    .map((text, i) => ({ text, top: i * 37 })) // stable, arbitrary-looking positions
    .sort(() => 0.5 - Math.random())

  it('reassembles reading order and every agent still reads its field correctly', () => {
    const lines = linesFromOcr(scrambledOcrLines)
    expect(lines).toEqual(L1_LINES) // reading order fully restored
    expect(detectTenderId(lines)).toBe('726879')
    expect(detectNitNoAndDate(lines).noticeNo).toBe('17/DB/EE/NizampetCircle58/CMC/2026-27')
    expect(detectNameOfWork(lines)).toContain('Laying of CC Road from Nayara Petrol Bunk')
    expect(detectL1Agency(lines)).toBe('NANDU CONSTRUCTIONS')
  })

  it('drops blank OCR detections rather than passing empty lines through', () => {
    const withBlanks = [{ text: 'Tender ID 726879', top: 0 }, { text: '   ', top: 5 }, { text: '', top: 10 }]
    expect(linesFromOcr(withBlanks)).toEqual(['Tender ID 726879'])
  })
})

describe('tender agents — work unchanged against a spreadsheet\'s rows (via linesFromExcelRows)', () => {
  // A plausible one-row-per-field tender comparison sheet: label and value
  // in adjacent cells of the same row, exactly like a label-anchored PDF
  // line reads once flattened.
  const rows: (string | number | null)[][] = [
    ['Tender ID', 726879],
    ['NIT No', '17/DB/EE/NizampetCircle58/CMC/2026-27', 'ITEM 5 Dated:18.08.2026'],
    ['Name of Work', 'Laying of CC Road from Nayara Petrol Bunk to Bhandari Layout'],
    ['Estimated Contract Value', 1521909.0],
    [null, '', 'Company Name', 'Amount', 'Rank'],
    ['NANDU CONSTRUCTIONS', 1521909.0, 'Less', 8.66, 1390111.68, 'L-1']
  ]

  it('flattens rows into lines and the label-anchored agents read them correctly', () => {
    const lines = linesFromExcelRows(rows)
    expect(lines[0]).toBe('Tender ID 726879')
    expect(detectTenderId(lines)).toBe('726879')
    expect(detectNitNoAndDate(lines).noticeNo).toBe('17/DB/EE/NizampetCircle58/CMC/2026-27')
    expect(detectNameOfWork(lines)).toContain('Laying of CC Road from Nayara Petrol Bunk')
  })

  it('skips empty cells and blank rows rather than emitting empty lines', () => {
    expect(linesFromExcelRows([[null, '', undefined], ['', 'Tender ID', '', 726879]])).toEqual(['Tender ID 726879'])
  })
})

// A second, structurally different office (Serilingampally/Ameenpur, portal
// login "ee-ptcu-ghmc") — found via a Downloads-folder scan, not the
// NIZAMPET-58 corpus. Its NIT No has no "/DB/EE/", "Circle-<n>", or "/C-<n>"
// at all ("Engg-21/CMC/AMPR-C-47/2026-27"), and its page reconstruction
// scrambles cells into a different order than every Nizampet/Gajularamaram
// sample seen so far. Only ONE real tender's worth of evidence backs this —
// far less than the 300+ files behind the Nizampet fixes — so this is
// explicitly a narrower, best-effort addition: the header fields below
// (Tender ID, NIT No, Name of Work, ECV) test cleanly against real data, but
// the price-bid table (L1 Agency, Tender Percentage, Contract Value, Name of
// the Agencies) is deliberately left unsupported — this same one sample's
// own 3 rows don't even agree on where the rank sits, which isn't something
// a regex should be fit to.
describe('tender agents — a second office with a structurally different NIT format (single-sample evidence)', () => {
  // Real Serilingampally/Ameenpur Commercial Evaluation page (Tender ID
  // 724743) — the exact pdf.js-reconstructed lines.
  const SERILINGAMPALLY_LINES = [
    'Insert',
    'Ho',
    '8/15/26, 2:38 PM',
    'Stage',
    'Selected Form',
    'Welcome to ee-ptcu-ghmc',
    'Profile',
    '| Training',
    'Manuals',
    '| Logout',
    'eGP Portal - Tenders',
    'Government of Telangana',
    'P',
    'Dashboard Tender Creation Tender Evaluation LOA',
    'Indent',
    'Management',
    'Commercial Evaluation',
    'Tender Details',
    'Current',
    'Enquiry/IFB/Tender Engg-21/CMC/AMPR-C-47/2026-27, Dated:',
    'Tender ID 724743',
    'Notice Number 10-08-2026 (Item No.3)',
    'Laying of CC Patches Sunway Opus, Sri vani nagar and Vediri Town Ship in Ward No.271, Ameenpur, Circle-47, Serilingampally',
    'Name of Work',
    'Zone, CMC (Rs.38.00 Lakhs) (2026-27)',
    'Works',
    'Tender Category Tender Evaluation Type Percentage',
    'Estimated Contract',
    '2986947.00',
    'OPEN - NCB',
    'Tender Type',
    'Value',
    'Bid Submission Start',
    'Bid Submission Closing',
    '14/08/2026 04:00 PM',
    '11/08/2026 10:30 AM',
    'Date',
    'Date & Time'
  ]

  it('reads Tender ID, NIT No, Name of Work and ECV from this office\'s differently-shaped header', () => {
    const r = runTenderAgents(SERILINGAMPALLY_LINES)
    expect(r.tenderId).toBe('724743')
    expect(r.noticeNo).toBe('Engg-21/CMC/AMPR-C-47/2026-27')
    expect(r.nameOfWork).toBe(
      'Laying of CC Patches Sunway Opus, Sri vani nagar and Vediri Town Ship in Ward No.271, Ameenpur, Circle-47, Serilingampally Zone, CMC (Rs.38.00 Lakhs) (2026-27)'
    )
    expect(r.nameOfWork).not.toMatch(/\bWorks$/) // the bare "Works" line must not get glued onto the end
    expect(r.ecvRupees).toBe(2986947)
    expect(r.bidStart).toBe('14/08/2026 04:00 PM')
    expect(r.bidClose).toBe('11/08/2026 10:30 AM')
  })

  it('does not false-positive-warn about this valid NIT format lacking "/DB/" or "/SE/"', () => {
    const r = runTenderAgents(SERILINGAMPALLY_LINES)
    expect(r.warnings.some((w) => w.includes('/DB/'))).toBe(false)
  })

  it('still surfaces the real ECV-vs-title disagreement on this office\'s own data', () => {
    // Title rounds to Rs.38.00 Lakhs (₹38,00,000); the actual ECV is
    // ₹29,86,947 — a genuine ~21% gap, same useful signal as the Nizampet
    // cross-check, now proven on a second, unrelated office's real data.
    const r = runTenderAgents(SERILINGAMPALLY_LINES)
    expect(r.warnings.some((w) => w.includes('disagree'))).toBe(true)
  })

  it('leaves the price-bid table fields unsupported for this office rather than guessing at its inconsistent row order', () => {
    const r = runTenderAgents(SERILINGAMPALLY_LINES)
    expect(r.l1Agency).toBeUndefined()
    expect(r.tenderPercentage).toBeUndefined()
    expect(r.contractRupees).toBeUndefined()
    expect(r.agencyNames).toEqual([])
  })
})

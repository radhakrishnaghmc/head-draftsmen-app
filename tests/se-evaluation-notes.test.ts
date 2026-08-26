import { describe, it, expect } from 'vitest'
import { buildAgencyApprovalHtml, buildBidEvaluationHtml, type AgencyApprovalData, type BidEvaluationData } from '../core/seEvaluationNotes'
import { convertHtmlToDocx } from '../core/htmlToDocx'

const baseRefs = {
  zoneAbbr: 'QBZ',
  workName: 'Test work',
  estimateLakhs: '50.00',
  asAuthority: 'Commissioner, CMC',
  asRefLine: 'Commissioner, CMC, dated:',
  tsNo: '12',
  tsDate: '01.01.2026',
  financialYear: '2026-27',
  nitNo: '15',
  nitDate: '27.07.2026',
  itemNo: '1',
  techBidOpenDate: '06.08.2026'
}

const bidder = { sno: '1', name: 'ABC Constructions', ecv: '5000000', pct: '(-)10', tcv: '4500000', rank: 'L-1' }

describe('Agency Approval / Bid Evaluation Ref block alignment', () => {
  // Real bug report: "Ref: 1) …" put point 1's number after the "Ref:"
  // label while points 2-5 started flush at the margin — a hanging-indent
  // <p> can only hang at one x position per paragraph, so with "Ref:" only
  // on item 1 its number sat further right than the others. Fixed by
  // putting every point in its own borderless-table row (label/number/text
  // columns), so all five numbers land in the same column regardless of
  // which row carries the "Ref:" label.
  it('gives every Ref point (1 through 5) the same number-column width, in Agency Approval', () => {
    const data: AgencyApprovalData = { ...baseRefs, zone: 'Quthbullapur', bidEvalApprovedDate: '06.08.2026', bidders: [bidder] }
    const html = buildAgencyApprovalHtml(data)
    const numCellWidths = [...html.matchAll(/width:26pt;[^"]*">(\d\))/g)].map((m) => m[1])
    expect(numCellWidths).toEqual(['1)', '2)', '3)', '4)', '5)'])
    // Only the first row carries the "Ref:" label.
    expect(html.match(/<b>Ref:<\/b>/g)?.length).toBe(1)
  })

  it('gives every Ref point (1 through 4) the same number-column width, in Bid Evaluation (no 5th item)', () => {
    const data: BidEvaluationData = { ...baseRefs, bidders: [bidder] }
    const html = buildBidEvaluationHtml(data)
    const numCellWidths = [...html.matchAll(/width:26pt;[^"]*">(\d\))/g)].map((m) => m[1])
    expect(numCellWidths).toEqual(['1)', '2)', '3)', '4)'])
    expect(html.match(/<b>Ref:<\/b>/g)?.length).toBe(1)
  })

  // Real regression caught here: an earlier version of this fix used
  // percentage column widths ('7%'/'6%'/'87%') to make the preview render
  // correctly (see above) — but html-to-docx's buildTableCellWidth only
  // parses pt/px/cm/inch width strings, so a percentage silently produced
  // an `undefined` width attribute that corrupted the OOXML and threw
  // ("InvalidCharacterError: Invalid XML name: @w"), breaking real document
  // generation outright — even though every HTML-shape assertion above
  // still passed. Only actually running the real conversion catches this
  // class of bug; keep this test even though it's slow.
  it('actually converts to a real .docx without throwing, in both documents', async () => {
    const agencyData: AgencyApprovalData = { ...baseRefs, zone: 'Quthbullapur', bidEvalApprovedDate: '06.08.2026', bidders: [bidder] }
    const bidEvalData: BidEvaluationData = { ...baseRefs, bidders: [bidder] }
    await expect(convertHtmlToDocx(buildAgencyApprovalHtml(agencyData))).resolves.toBeInstanceOf(Buffer)
    await expect(convertHtmlToDocx(buildBidEvaluationHtml(bidEvalData))).resolves.toBeInstanceOf(Buffer)
  })
})

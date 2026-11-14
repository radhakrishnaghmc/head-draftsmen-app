import { describe, it, expect } from 'vitest'
import { buildBidEvaluationHtml, type AgencyApprovalData, type BidEvaluationData } from '../core/seEvaluationNotes'
import { buildAgencyApprovalDocxB64, buildAgencyApprovalDocxXml, agencyApprovalScale } from '../core/seAgencyApprovalDocx'
import PizZip from 'pizzip'
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
  it('gives every Ref point (1 through 4) the same number-column width, in Bid Evaluation (no 5th item)', () => {
    const data: BidEvaluationData = { ...baseRefs, bidders: [bidder] }
    const html = buildBidEvaluationHtml(data)
    const numCellWidths = [...html.matchAll(/width:26pt;[^"]*"><p[^>]*>(\d\))/g)].map((m) => m[1])
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
    const bidEvalData: BidEvaluationData = { ...baseRefs, bidders: [bidder] }
    await expect(convertHtmlToDocx(buildBidEvaluationHtml(bidEvalData))).resolves.toBeInstanceOf(Buffer)
  })
})

describe('Ref block table grid', () => {
  it('sizes the grid columns to the cell widths, not equal thirds', async () => {
    const { convertHtmlToDocx } = await import('../core/htmlToDocx')
    const { buildBidEvaluationHtml } = await import('../core/seEvaluationNotes')
    const PizZip = (await import('pizzip')).default
    const d: any = { zoneAbbr: 'QBZ', corp: 'CMC', workName: 'W', estimateLakhs: '1', asAuthority: 'x', asRefLine: 'x', tsNo: '1', tsDate: 'd', financialYear: '2026-27', nitNo: 'n', nitDate: 'd', itemNo: '1', techBidOpenDate: 'd', bidders: [] }
    const xml = new PizZip(await convertHtmlToDocx(buildBidEvaluationHtml(d))).file('word/document.xml')!.asText()
    const grid = /<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/.exec(xml)![1]
    expect([...grid.matchAll(/w:w="(\d+)"/g)].map((m) => m[1])).toEqual(['600', '520', '7520'])
  })
})

describe('Agency Approval note (strict A4)', () => {
  const mk = (n: number): AgencyApprovalData => ({
    ...baseRefs,
    zone: 'Quthbullapur',
    bidEvalApprovedDate: '06.08.2026',
    bidders: Array.from({ length: n }, (_, i) => ({ ...bidder, sno: String(i + 1), name: `Agency ${i + 1}`, rank: `L-${i + 1}` }))
  })
  const xmlOf = (d: AgencyApprovalData) => new PizZip(buildAgencyApprovalDocxB64(d), { base64: true }).file('word/document.xml')!.asText()

  it('is an A4 page whose tables never exceed the text width', () => {
    const xml = xmlOf(mk(8))
    expect(xml).toContain('<w:pgSz w:w="11906" w:h="16838"')
    const m = /<w:pgMar[^>]*w:right="(\d+)"[^>]*w:left="(\d+)"/.exec(xml)!
    const text = 11906 - Number(m[1]) - Number(m[2])
    for (const t of xml.matchAll(/<w:tblW w:w="(\d+)"/g)) expect(Number(t[1])).toBeLessThanOrEqual(text)
  })

  it('shrinks table text as more agencies bid (11pt for 2, 10pt for 5, 9pt for 8)', () => {
    expect(agencyApprovalScale(2).size).toBe(22)
    expect(agencyApprovalScale(5).size).toBe(20)
    expect(agencyApprovalScale(8).size).toBe(18)
    expect(agencyApprovalScale(20).size).toBeLessThan(agencyApprovalScale(8).size)
  })

  it('lists every bidder, names the L-1 and signs off AE / Dy.EE / SE (M)', () => {
    const xml = buildAgencyApprovalDocxXml(mk(3))
    for (const i of [1, 2, 3]) expect(xml).toContain(`Agency ${i}`)
    expect(xml).toContain('is the lowest bidder and quoted @ (-) 10 % Less than the ECV')
    expect(xml).toContain('found that (03) bidders were eligible')
    expect(xml).toContain('SE (M), QBZ')
  })

  it('names the chosen corporation instead of CMC', () => {
    const xml = buildAgencyApprovalDocxXml({ ...mk(1), corp: 'MMC', asRefLine: 'Commissioner, MMC, dated:' })
    expect(xml).toContain('MMC')
    expect(xml).not.toMatch(/\bCMC\b/)
  })
})

import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { parseTenderEvaluation } from '../core/tenderEvaluationPdf'
import { parseParticipatingBidders } from '../core/viewBiddersPdf'
import { buildEvaluationSheet, columnLetter, EVALUATION_CRITERIA } from '../core/evaluationSheet'

// The exact lines pdf.js reconstructs from the portal's "View Bidders"
// (Supplier List) page for Tender 722231 — including the wrapped bidder name
// (BOBBA RAVI CHANDRA CIVIL … CONTRACTOR) whose number row lands between its two
// name lines.
const VIEW_BIDDERS = [
  'Welcome to ee-grrc-ghmc Profile | Training Manuals | Logout',
  'Evaluation Stages',
  'Current Tender Details',
  'Enquiry/IFB/Tender',
  '16/DB/EE/Gajularamaram Circle-',
  '722231',
  'Tender ID',
  '57/CMC/2026-27,Item NO.3,Dt:25.07.2026',
  'Notice Number',
  'Name of Work Laying of CC road In Internal lanes of Sai Gramam Colony in ward no. 277, Mahadevapuram, Gajularamaram Circle-57, CMC',
  'Tender Category Works Tender Evaluation Type Percentage',
  'Estimated Contract',
  'Tender Type OPEN - NCB 1723159.00',
  'Value',
  'Bid Submission Start Bid Submission Closing',
  '28/07/2026 08:00 PM 04/08/2026 02:30 PM',
  'Date & Time Date',
  'Supplier List /Commercial Stage',
  'Company Name Action',
  'K S REDDY AND COMPANY 1340371-0 04/08/2026 02:19 PM 49.43.221.124',
  'G GANESH 1340239-0 04/08/2026 01:01 PM 106.222.229.181',
  'PATLOLLA SRINIVAS REDDY 1339560-2 04/08/2026 11:29 AM 49.204.97.90',
  'BOBBA RAVI CHANDRA CIVIL',
  '1339995-1 04/08/2026 12:36 AM 106.222.233.188',
  'CONTRACTOR',
  'GK CONSTRUCTIONS 1340339-0 04/08/2026 01:49 PM 106.222.233.213',
  'Back Bulk Download'
]

describe('parseParticipatingBidders', () => {
  it('reads every bidder, reassembling the wrapped name', () => {
    const names = parseParticipatingBidders(VIEW_BIDDERS).map((b) => b.name)
    expect(names).toEqual([
      'K S REDDY AND COMPANY',
      'G GANESH',
      'PATLOLLA SRINIVAS REDDY',
      'BOBBA RAVI CHANDRA CIVIL CONTRACTOR',
      'GK CONSTRUCTIONS'
    ])
  })

  it('captures the bid number and submission time', () => {
    const first = parseParticipatingBidders(VIEW_BIDDERS)[0]
    expect(first.bidNumber).toBe('1340371-0')
    expect(first.submittedAt).toBe('04/08/2026 02:19 PM')
  })

  it('the shared header parse still reads tender id and ECV from the same page', () => {
    const ev = parseTenderEvaluation(VIEW_BIDDERS)
    expect(ev.tenderId).toBe('722231')
    expect(ev.ecvRupees).toBe(1723159)
  })

  it('returns [] for a page with no supplier rows', () => {
    expect(parseParticipatingBidders(['Current Tender Details', 'Company Name Action'])).toEqual([])
  })
})

describe('columnLetter', () => {
  it('maps 1-based indexes to Excel column letters', () => {
    expect([1, 2, 3, 26, 27, 28].map(columnLetter)).toEqual(['A', 'B', 'C', 'Z', 'AA', 'AB'])
  })
})

async function load(buf: Buffer): Promise<ExcelJS.Worksheet> {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf as unknown as ArrayBuffer)
  return wb.worksheets[0]
}

describe('buildEvaluationSheet', () => {
  const input = {
    nitLine: 'NIT No: 16/DB/EE/Gajularamaram Circle-57/CMC/2026-27,Dt:25.07.2026,Tender ID:722231',
    workLine: 'Name of the work: Laying of CC road. ecv: 1723159.00',
    bidders: ['K S REDDY AND COMPANY', 'G GANESH', 'GK CONSTRUCTIONS'],
    ecv: 1723159,
    signature: 'Executive Engineer,\nGajularamaram circle,CMC'
  }

  it('lays out the header, one column per bidder and the ECV', async () => {
    const ws = await load(await buildEvaluationSheet(input))
    expect(ws.getCell('A1').value).toBe('TENDER DETAILS')
    expect(ws.getCell('A2').value).toContain('Tender ID:722231')
    expect(ws.getCell('C4').value).toBe('K S REDDY AND COMPANY')
    expect(ws.getCell('D4').value).toBe('G GANESH')
    expect(ws.getCell('E4').value).toBe('GK CONSTRUCTIONS')
    expect(ws.getCell('C20').value).toBe(1723159)
    expect(ws.getCell('E20').value).toBe(1723159)
    // Signature sits two rows below the surplus row (26), under the last bidder column (E).
    expect(ws.getCell('E26').value).toBe('Executive Engineer,\nGajularamaram circle,CMC')
  })

  it('wires the bid-capacity and surplus formulas per bidder column', async () => {
    const ws = await load(await buildEvaluationSheet(input))
    expect((ws.getCell('C19').value as { formula: string }).formula).toBe('(3*C16*C17)-C18')
    expect((ws.getCell('E19').value as { formula: string }).formula).toBe('(3*E16*E17)-E18')
    expect((ws.getCell('D22').value as { formula: string }).formula).toBe('D19-D20')
  })

  it('prints all eleven criteria and leaves the data inputs (incl. Remarks) blank', async () => {
    const ws = await load(await buildEvaluationSheet(input))
    expect(String(ws.getCell('B5').value)).toBe(EVALUATION_CRITERIA[0])
    expect(String(ws.getCell('B15').value)).toBe(EVALUATION_CRITERIA[10])
    expect(ws.getCell('B16').value).toContain('Turnover')
    expect(ws.getCell('C16').value).toBeNull() // turnover (A) blank
    expect(ws.getCell('C17').value).toBeNull() // N blank
    expect(ws.getCell('C18').value).toBeNull() // B blank
    expect(String(ws.getCell('B21').value)).toBe('Remarks')
    expect(ws.getCell('C21').value).toBeNull() // Remarks blank under each bidder
  })

  it('sizes the Name of Work row to fit long text instead of clipping it', async () => {
    const shortWs = await load(await buildEvaluationSheet(input))
    const shortHeight = shortWs.getRow(3).height ?? 0

    const longWs = await load(
      await buildEvaluationSheet({
        ...input,
        workLine: 'Name of the work: ' + 'A very long work name that wraps across several lines '.repeat(5) + 'ecv: 1723159.00'
      })
    )
    const longHeight = longWs.getRow(3).height ?? 0

    expect(shortHeight).toBeGreaterThan(0)
    expect(longHeight).toBeGreaterThan(shortHeight)
  })
})

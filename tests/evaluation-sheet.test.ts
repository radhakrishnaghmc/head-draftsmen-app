import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { parseTenderEvaluation } from '../core/tenderEvaluationPdf'
import { parseParticipatingBidders } from '../core/viewBiddersPdf'
import { applyReservationConditions, buildEvaluationSheet, columnLetter, EVALUATION_CRITERIA } from '../core/evaluationSheet'

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
    workLine: 'Name of the work: Laying of CC road (Reserved for SC). ecv: 1723159.00',
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

  it('replaces the standard criteria with the tender\'s own conditions and shifts the block down', async () => {
    const base = {
      nitLine: 'NIT',
      workLine: 'Work',
      bidders: ['A', 'B'],
      ecv: 600000,
      estimate: 600000,
      signature: 'EE'
    }
    const conditions = ['Registration Certification', 'EMD', 'Transaction Details']
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load((await buildEvaluationSheet({ ...base, conditions })) as unknown as ArrayBuffer)
    const ws = wb.worksheets[0]
    expect(ws.getCell(5, 2).value).toBe('Registration Certification')
    expect(ws.getCell(7, 2).value).toBe('Transaction Details')
    // Bid-capacity block follows the three conditions: A at row 8, 3AN-B at row 11.
    expect(ws.getCell(8, 1).value).toBe(4)
    expect(String(ws.getCell(8, 2).value)).toContain('Turnover')
    const f = ws.getCell(11, 3).value as { formula: string }
    expect(f.formula).toBe('(3*C8*C9)-C10')
    expect(ws.getCell(12, 3).value).toBe(600000)
  })

  it('drops the 3AN-B block for works at or below 5 lakhs, keeps it above', async () => {
    const base = { nitLine: 'NIT', workLine: 'Work', bidders: ['A'], signature: 'EE', conditions: ['Registration', 'EMD'] }
    const labels = async (estimate: number | null) => {
      const wb = new ExcelJS.Workbook()
      await wb.xlsx.load((await buildEvaluationSheet({ ...base, ecv: 450000, estimate })) as unknown as ArrayBuffer)
      const ws = wb.worksheets[0]
      const out: string[] = []
      ws.eachRow((row) => out.push(String(row.getCell(2).value ?? '')))
      return out.join('|')
    }
    const small = await labels(400260)
    expect(small).not.toMatch(/3AN-B|Turnover/)
    expect(small).toMatch(/ECV\|Remarks/)
    expect(await labels(500000)).not.toMatch(/3AN-B/)
    expect(await labels(500001)).toMatch(/3AN-B/)
    // No estimate on the Works List: keep the rows, whatever the ECV is.
    expect(await labels(null)).toMatch(/3AN-B/)
  })

  it('omits the bid-capacity standard criteria for small works', async () => {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(
      (await buildEvaluationSheet({ nitLine: 'N', workLine: 'W', bidders: ['A'], ecv: 900000, estimate: 300000, signature: 'EE' })) as unknown as ArrayBuffer
    )
    const text = wb.worksheets[0].getColumn(2).values.join('|')
    expect(text).not.toMatch(/3AN-B|existing commitments|Maximum value of Civil/)
    expect(text).toMatch(/GST Registration/)
  })
})

describe('reservation conditions', () => {
  const base = {
    nitLine: 'NIT No: 16/DB/EE/Gajularamaram Circle-57/CMC/2026-27,Dt:25.07.2026,Tender ID:722231',
    bidders: ['A', 'B'],
    ecv: 1723159,
    signature: 'Executive Engineer,\nGajularamaram circle,CMC'
  }
  const slNos = (ws: ExcelJS.Worksheet, rows: number) => Array.from({ length: rows }, (_, i) => ws.getCell(5 + i, 1).value)

  async function sheet(workLine: string, conditions?: string[]) {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load((await buildEvaluationSheet({ ...base, workLine, conditions })) as unknown as ArrayBuffer)
    return wb.worksheets[0]
  }
  const texts = (ws: ExcelJS.Worksheet, rows: number) => Array.from({ length: rows }, (_, i) => String(ws.getCell(5 + i, 2).value))

  it('keeps the caste-certificate and one-work conditions for a reserved work, in any category', async () => {
    for (const cat of ['SC', 'ST', 'Waddera', 'Vaddera']) {
      const ws = await sheet(`Name of the work: Laying of CC road (Reserved for ${cat}). ecv: 1.00`)
      const t = texts(ws, 11)
      expect(t.some((x) => /^Caste Certificate/.test(x))).toBe(true)
      expect(t.some((x) => /^For reservation work only one/.test(x))).toBe(true)
    }
  })

  it('drops both for an open work and keeps the serial numbers 1..9 unbroken', async () => {
    const ws = await sheet('Name of the work: Laying of CC road. ecv: 1.00')
    const t = texts(ws, 9)
    expect(t.some((x) => /caste certificate|reservation work/i.test(x))).toBe(false)
    expect(slNos(ws, 9)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9])
    // The bid-capacity rows continue the sequence straight after the last condition.
    expect(ws.getCell(14, 1).value).toBe(10)
  })

  it('applies to a portal-fetched list too: removes them when open, adds them when reserved and missing', () => {
    const portal = ['Registration', 'GST', 'Caste Certificate to be submitted by the individuals', 'For reservation work only one work should be allotted', 'Self-declaration of genuinity']
    expect(applyReservationConditions(portal, false)).toEqual(['Registration', 'GST', 'Self-declaration of genuinity'])
    const noReservation = ['Registration', 'GST', 'Self-declaration of genuinity']
    const withThem = applyReservationConditions(noReservation, true)
    expect(withThem).toHaveLength(5)
    expect(withThem[4]).toBe('Self-declaration of genuinity')
    expect(withThem[2]).toMatch(/^Caste Certificate/)
  })
})

describe('estimate below 5 lakhs', () => {
  const base = {
    nitLine: 'NIT No: 1',
    workLine: 'Name of the work: Repairs. ecv: 400000.00',
    bidders: ['A'],
    ecv: 400000,
    signature: 'EE'
  }
  async function labels(input: object) {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load((await buildEvaluationSheet({ ...base, ...input })) as unknown as ArrayBuffer)
    const ws = wb.worksheets[0]
    const rows: { sl: unknown; text: string }[] = []
    ws.eachRow((r, n) => n >= 5 && rows.push({ sl: r.getCell(1).value, text: String(r.getCell(2).value) }))
    return rows
  }

  it('drops the bid-capacity conditions and the A / N / B / 3AN-B rows, keeping serials unbroken', async () => {
    const rows = await labels({ estimate: 500000 })
    const t = rows.map((r) => r.text).join('|')
    expect(t).not.toMatch(/existing commitments|administrative sanction|Maximum value of Civil|Turnover in Civil|3AN-B/i)
    expect(rows.map((r) => r.text)).not.toContain('N')
    expect(rows.map((r) => r.text)).not.toContain('B')
    const numbered = rows.filter((r) => r.sl != null).map((r) => r.sl)
    expect(numbered).toEqual(numbered.map((_, i) => i + 1)) // Remarks carries no serial
  })

  it('also drops them from a portal-fetched list, matched by wording', async () => {
    const rows = await labels({
      estimate: 500000,
      conditions: ['GST', 'Value of existing commitments and ongoing works', 'The administrative sanction amount above 5.00 lakhs will be evaluated based on 3AN-B', 'Maximum value of Civil Engineering works executed in any one year', 'Self-declaration']
    })
    expect(rows.slice(0, 2).map((r) => r.text)).toEqual(['GST', 'Self-declaration'])
  })

  it('keeps them when the estimate is above 5 lakhs', async () => {
    const t = (await labels({ estimate: 900000 })).map((r) => r.text).join('|')
    expect(t).toMatch(/existing commitments/)
    expect(t).toMatch(/3AN-B/)
  })
})

describe('conditionsFinal', () => {
  it('prints an edited list exactly, with no reservation or small-work filtering', async () => {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(
      (await buildEvaluationSheet({
        nitLine: 'NIT',
        workLine: 'Name of the work: Repairs',
        bidders: ['A'],
        ecv: 1,
        estimate: 100,
        signature: 'EE',
        conditions: ['Only this', 'Caste Certificate to be submitted by me'],
        conditionsFinal: true
      })) as unknown as ArrayBuffer
    )
    const ws = wb.worksheets[0]
    expect([ws.getCell('B5').value, ws.getCell('B6').value]).toEqual(['Only this', 'Caste Certificate to be submitted by me'])
    expect(ws.getCell('A6').value).toBe(2)
  })
})

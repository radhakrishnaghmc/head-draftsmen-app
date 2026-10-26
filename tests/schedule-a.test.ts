import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import * as ExcelJS from 'exceljs'
import { buildScheduleARows, metaFromWorksRow, type ScheduleAItem } from '../core/scheduleA'
import { fillScheduleATemplate, fillSeScheduleATemplate } from '../core/scheduleATemplate'
import { scheduleAWorkbookToHtml } from '../core/scheduleAHtml'

const TEMPLATE_PATH = resolve(__dirname, '../resources/schedule-a-template.xlsx')
const SE_TEMPLATE_PATH = resolve(__dirname, '../resources/se-schedule-a-template.xlsx')

const items: ScheduleAItem[] = [
  { itemNo: '1', quantity: '10', description: 'Earth work', rate: '100', units: 'Cum', amount: '1000' }
]

describe('buildScheduleARows', () => {
  it('shows the meta amounts Indian-formatted with a "/-" suffix, not the raw Lakhs figure', () => {
    const meta = metaFromWorksRow({
      'Name of the work': 'Road from A to B',
      'Amount of estimate': '45',
      'ECV': '4200000',
      'Tender Percentage': '18',
      'Name of the Agency': 'ABC Constructions'
    })
    const rows = buildScheduleARows(items, meta)
    const estimateRow = rows.find((r) => r[0] === 'Estimate Amount: Rs.')
    const ecvRow = rows.find((r) => r[0] === 'ECV Amount: Rs.')
    const contractRow = rows.find((r) => r[0] === 'Contract Amount: Rs.')
    expect(estimateRow?.[2]).toBe('45,00,000/-')
    expect(ecvRow?.[2]).toBe('42,00,000/-')
    expect(contractRow?.[2]).toBe('34,44,000/-') // 42,00,000 * (1 - 0.18)
  })

  it('falls back to the item total, Indian-formatted, when there is no Works List match', () => {
    const rows = buildScheduleARows(items) // no meta at all
    const estimateRow = rows.find((r) => r[0] === 'Estimate Amount: Rs.')
    expect(estimateRow?.[2]).toBe('1,000/-')
  })

  it('leaves ECV Amount blank (never the item total) for a matched row whose ECV is blank', () => {
    const meta = metaFromWorksRow({ 'Name of the work': 'Road from A to B', 'Amount of estimate': '45', 'ECV': '' })
    const rows = buildScheduleARows(items, meta)
    const ecvRow = rows.find((r) => r[0] === 'ECV Amount: Rs.')
    expect(ecvRow?.[2]).toBe('')
  })

  it('leaves Contract Amount blank when Tender Percentage is not on the Works List row', () => {
    const meta = metaFromWorksRow({
      'Name of the work': 'Road from A to B',
      'Amount of estimate': '45',
      'ECV': '4200000'
    })
    const rows = buildScheduleARows(items, meta)
    const contractRow = rows.find((r) => r[0] === 'Contract Amount: Rs.')
    expect(contractRow?.[2]).toBe('')
  })
})

describe('fillScheduleATemplate', () => {
  it('writes the Indian-formatted estimate/ECV/contract amounts into the template, not the raw Lakhs figure', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const meta = metaFromWorksRow({
      'Name of the work': 'Road from A to B',
      'Amount of estimate': '45',
      'ECV': '4200000',
      'Tender Percentage': '18',
      'Name of the Agency': 'ABC Constructions'
    })
    const out = await fillScheduleATemplate(buffer, items, meta)

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]

    const values: string[] = []
    for (let r = 1; r <= 20; r++) values.push(String(ws.getCell(r, 3).value ?? ''))
    // "Rs." sits in the value cell (not the label) so the three amounts line up in one column.
    expect(values).toContain('Rs. 45,00,000/-')
    expect(values).toContain('Rs. 42,00,000/-')
    expect(values).toContain('Rs. 34,44,000/-')
    expect(String(ws.getCell(12, 1).value)).toBe('Estimate Amount:')
    // The raw Lakhs figures must not appear as bare numbers anywhere in that column.
    expect(values).not.toContain('45')
    expect(values).not.toContain('42')
  }, 30000)

  it("replaces the template's sample circle/zone with the work's own circle/zone", async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const meta = metaFromWorksRow({
      'Name of the work': 'Road from A to B',
      Circle: 'Nizampet',
      'Circle number': '58',
      Zone: 'Quthbullapur',
      'Amount of estimate': '45'
    })
    expect(meta.circle).toBe('Nizampet Circle-58')
    const out = await fillScheduleATemplate(buffer, items, meta)

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]
    let allText = ''
    ws.eachRow((row) => row.eachCell((c) => { allText += ' ' + (typeof c.value === 'string' ? c.value : '') }))

    // The sample circle from the template must be gone, replaced by this work's.
    expect(allText).not.toContain('Gajularamaram Circle-57')
    expect(allText).toContain('Nizampet Circle-58,CMC')
  }, 30000)

  it('leaves the office text alone when no circle/zone is supplied', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = await fillScheduleATemplate(buffer, items, { nameOfWork: 'X' })
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]
    let allText = ''
    ws.eachRow((row) => row.eachCell((c) => { allText += ' ' + (typeof c.value === 'string' ? c.value : '') }))
    expect(allText).toContain('Gajularamaram Circle-57')
  }, 30000)

  it('auto-sizes each item row to fit its wrapping description, with wrapText on', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const longDesc =
      'Manufacture as per BIS:12592 (part 1 & 2) supply and delivery of heavy-duty manhole covers and frames, including cost and conveyance of all materials, all charges etc. complete for the finished item of work as per the directions of the Engineer-In-Charge.'
    const twoItems: ScheduleAItem[] = [
      { itemNo: '1', quantity: '4', description: 'Short item', rate: '100', units: 'Each', amount: '400' },
      { itemNo: '2', quantity: '4', description: longDesc, rate: '2716', units: '', amount: '10864' }
    ]
    const out = await fillScheduleATemplate(buffer, twoItems, { nameOfWork: 'X' })
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]
    let shortH = 0
    let longH = 0
    let longWraps = false
    ws.eachRow((row, r) => {
      const d = String(ws.getCell(r, 3).value ?? '')
      if (d === 'Short item') shortH = row.height ?? 0
      if (d.startsWith('Manufacture as per BIS')) {
        longH = row.height ?? 0
        longWraps = !!ws.getCell(r, 3).alignment?.wrapText
      }
    })
    expect(shortH).toBeGreaterThan(0)
    expect(longH).toBeGreaterThan(shortH) // the long description forced a taller row
    expect(longWraps).toBe(true)
  }, 30000)
})

describe('fillSeScheduleATemplate', () => {
  // A Zone-level (SE) office's Schedule A must sign off as Superintending
  // Engineer with the Zone/Corporation — never the Executive Engineer/Circle
  // wording the EE template carries. Real bug: the standalone "Save Schedule
  // A" button already branched on this correctly, but the Agreement tab's
  // "Download all documents" bundle (electron/main.ts buildScheduleABuffer)
  // always called fillScheduleATemplate (the EE one) regardless of office —
  // an SE office's downloaded Schedule A showed "Executive Engineer" and a
  // Circle name instead of "Superintending Engineer" and the Zone.
  it('signs off as Superintending Engineer with the Zone and Corporation, never Executive Engineer/Circle', async () => {
    const buffer = readFileSync(SE_TEMPLATE_PATH)
    const meta = { nameOfWork: 'Road from A to B', zone: 'Quthbullapur', corporation: 'CMC' }
    const out = await fillSeScheduleATemplate(buffer, items, meta)

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]
    let allText = ''
    ws.eachRow((row) => row.eachCell((c) => { allText += ' ' + (typeof c.value === 'string' ? c.value : '') }))

    expect(allText).toContain('Superintending Engineer')
    expect(allText).toContain('Quthbullapur Zone, CMC')
    expect(allText).not.toContain('{{Zone}}')
    expect(allText).not.toContain('{{Corp}}')
  }, 30000)

  it('leaves the {{Zone}}/{{Corp}} placeholders alone when neither is supplied', async () => {
    const buffer = readFileSync(SE_TEMPLATE_PATH)
    const out = await fillSeScheduleATemplate(buffer, items, { nameOfWork: 'X' })
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]
    let allText = ''
    ws.eachRow((row) => row.eachCell((c) => { allText += ' ' + (typeof c.value === 'string' ? c.value : '') }))
    expect(allText).toContain('{{Zone}}')
    expect(allText).toContain('Superintending Engineer')
  }, 30000)
})

describe('Schedule A print coverage (through the Executive Engineer signature)', () => {
  const many: ScheduleAItem[] = Array.from({ length: 30 }, (_, i) => ({
    itemNo: String(i + 1),
    quantity: '10',
    description: `Item ${i + 1}`,
    rate: '100',
    units: 'Cum',
    amount: '1000'
  }))

  for (const [label, tpl, fill, title] of [
    ['EE', TEMPLATE_PATH, fillScheduleATemplate, 'Executive Engineer'],
    ['SE', SE_TEMPLATE_PATH, fillSeScheduleATemplate, 'Superintending Engineer']
  ] as const) {
    it(`${label}: the print area ends at the signature block's last row, however many items were inserted`, async () => {
      const out = await fill(readFileSync(tpl), many, { nameOfWork: 'X' })
      const wb = new ExcelJS.Workbook()
      await wb.xlsx.load(out as unknown as ArrayBuffer)
      const ws = wb.worksheets[0]
      let titleRow = 0
      let lastRow = 0
      ws.eachRow((row, i) => {
        row.eachCell((c) => {
          if (c.value === title) titleRow = i
        })
        lastRow = i
      })
      expect(titleRow).toBeGreaterThan(29) // the signature really did move below the template's old print area
      expect(ws.pageSetup.printArea).toBe(`A1:F${lastRow}`)
      expect(lastRow).toBeGreaterThan(titleRow) // the circle / zone line below the title is covered too
    }, 30000)

    it(`${label}: the PDF page keeps the signature block and evaluates the amounts`, async () => {
      const out = await fill(readFileSync(tpl), many, { nameOfWork: 'X' })
      const { html } = await scheduleAWorkbookToHtml(out)
      expect(html).toContain(title)
      expect(html).toContain('1000.00') // 10 x 100 per item, in the sheet's own 0.00 format
      expect(html).toContain('30000.00') // SUM of the 30 amounts
      expect(html).not.toMatch(/style="[^"]*"[^>]*"[^>]*>/) // no quote inside a style attribute closes it early
      expect(html).not.toContain('<thead>') // the item header is not repeated on the next page
      expect(html).not.toContain('#REF!') // the template's scratch formulas never print
    }, 30000)
  }
})

describe('Schedule A "In Figures" row', () => {
  it('spells out the computed tender-adjusted total', async () => {
    const out = await fillScheduleATemplate(readFileSync(TEMPLATE_PATH), [
      { itemNo: '1', quantity: '10', description: 'A', rate: '1000', units: 'Cum', amount: '0' }
    ], { tenderPercentage: '10' } as never)
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load(out as unknown as ArrayBuffer)
    let label = ''
    wb.worksheets[0].eachRow((row) => {
      const v = row.getCell(1).value
      if (typeof v === 'string' && /^in figures/i.test(v)) label = v
    })
    expect(label).toBe('In Figures (Nine Thousand Rupees Only)') // 10,000 less 10%
  }, 30000)
})

import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { buildDetailedEstimateWorkbook } from '../core/estimateTemplate'
import type { EstimateWorkItem } from '../core/estimateExtract'

async function loadSheet(buffer: Buffer) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer as unknown as ArrayBuffer)
  return wb.worksheets[0]
}

function cellText(ws: ExcelJS.Worksheet, row: number, col: number): string {
  const v = ws.getCell(row, col).value
  if (v && typeof v === 'object' && 'formula' in v) return `=${(v as { formula: string }).formula}`
  return v === null || v === undefined ? '' : String(v)
}

const items: EstimateWorkItem[] = [
  { description: 'Earthwork excavation', quantity: '100.00', rate: '200.00', unit: 'Cum' },
  { description: 'Providing wet mix macadam', quantity: '50.00', rate: '2000.00', unit: 'Cum' }
]

describe('buildDetailedEstimateWorkbook', () => {
  it('writes the letterhead, title, name of work, and a live Estimate Amount formula off the Grand Total', async () => {
    const ws = await loadSheet(
      await buildDetailedEstimateWorkbook(items, { zone: 'Quthbullapur', circle: 'Nizampet Circle-58', workName: 'Laying of cc road' })
    )
    expect(cellText(ws, 1, 1)).toBe('Nizampet Circle-58 :: Quthbullapur ZONE')
    expect(cellText(ws, 2, 1)).toBe('Detailed and Abstract - ESTIMATE')
    expect(cellText(ws, 3, 1)).toBe('Name of the work: Laying of cc road')
    expect(cellText(ws, 4, 6)).toBe('Estimate Amount Rs. ')
    expect(cellText(ws, 4, 11)).toBe('=ROUND(L23/100000,2)') // 2 items x 3 rows each -> Grand Total lands on row 23
    expect(cellText(ws, 4, 12)).toBe('lakhs')
  })

  it('writes the standard column headers', async () => {
    const ws = await loadSheet(await buildDetailedEstimateWorkbook(items))
    expect(cellText(ws, 5, 1)).toContain('Sl')
    expect(cellText(ws, 5, 2)).toBe('Description of work')
    expect(cellText(ws, 5, 9)).toBe('Qty')
    expect(cellText(ws, 5, 10)).toBe('Rate')
    expect(cellText(ws, 5, 11)).toBe('Per')
    expect(cellText(ws, 5, 12)).toBe('Amount')
  })

  it('writes each item across 3 rows — lead (Sl.No + Description only), a blank dimension row, then a summary row with Qty/Rate/Per/Amount', async () => {
    const ws = await loadSheet(await buildDetailedEstimateWorkbook(items))
    // Item 1: lead=6, dimension=7, summary=8.
    expect(cellText(ws, 6, 1)).toBe('1')
    expect(cellText(ws, 6, 2)).toBe('Earthwork excavation')
    // Rate/Per/Amount (columns 10-12) sit outside the lead row's B:I merge —
    // confirming Qty/Rate/Amount genuinely aren't on the same row as the description.
    expect(cellText(ws, 6, 12)).toBe('')

    expect(cellText(ws, 7, 2)).toBe('') // dimension row: no L/B/D data to show, left blank
    expect(cellText(ws, 7, 9)).toBe('')

    expect(cellText(ws, 8, 1)).toBe('') // Sl.No/Description do not repeat on the summary row
    expect(cellText(ws, 8, 9)).toBe('100')
    expect(cellText(ws, 8, 10)).toBe('200')
    expect(cellText(ws, 8, 11)).toBe('Cum')
    expect(cellText(ws, 8, 12)).toBe('=ROUND(I8*J8,0)')

    // Item 2: lead=9, dimension=10, summary=11.
    expect(cellText(ws, 9, 1)).toBe('2')
    expect(cellText(ws, 9, 2)).toBe('Providing wet mix macadam')
    expect(cellText(ws, 11, 9)).toBe('50')
  })

  it('fills the dimension row with No\'s/L/B/D when the source estimate provides them', async () => {
    const itemsWithDims: EstimateWorkItem[] = [
      { description: 'Providing wet mix macadam', quantity: '30.00', rate: '2000.00', unit: 'Cum', nos: '5', l: '10', b: '2', d: '0.3' }
    ]
    const ws = await loadSheet(await buildDetailedEstimateWorkbook(itemsWithDims))
    // dimension row = 7 (lead=6, dimension=7, summary=8), No's=col3, L=col6, B=col7, D=col8
    expect(cellText(ws, 7, 3)).toBe('5')
    expect(cellText(ws, 7, 6)).toBe('10')
    expect(cellText(ws, 7, 7)).toBe('2')
    expect(cellText(ws, 7, 8)).toBe('0.3')
  })

  it('builds the standard surcharge cascade exactly matching the real template’s own formulas', async () => {
    const ws = await loadSheet(await buildDetailedEstimateWorkbook(items))
    // firstItemRow=6, 2 items x 3 rows -> lastItemRow=11, subTotal1Row=12
    expect(cellText(ws, 12, 2)).toBe('Sub Total ')
    expect(cellText(ws, 12, 12)).toBe('=SUM(L6:L11)')
    expect(cellText(ws, 13, 2)).toBe('LC 1.00%')
    expect(cellText(ws, 13, 12)).toBe('=ROUND(L12*1%,0)')
    expect(cellText(ws, 14, 12)).toBe('=ROUND(L12*0.1%,0)') // NAC
    expect(cellText(ws, 15, 12)).toBe('=ROUND(L12*0.35%,0)') // TPQC
    // Seigniorage: 2% of the Sub Total (the ECV) — a live formula, not manual entry.
    expect(cellText(ws, 16, 2)).toBe('Seigniorage charges ')
    expect(cellText(ws, 16, 12)).toBe('=ROUND(L12*2%,0)')
    expect(cellText(ws, 17, 12)).toBe('=ROUND(30%*L16,0)') // DMFT
    expect(cellText(ws, 18, 12)).toBe('=ROUND(L16*2%,0)') // SMET
    expect(cellText(ws, 19, 12)).toBe('=ROUND(L16*0.8,0)') // Permit Fee
    expect(cellText(ws, 20, 2)).toBe('Sub Total ')
    expect(cellText(ws, 20, 12)).toBe('=SUM(L12:L19)')
    expect(cellText(ws, 21, 12)).toBe('=ROUND(L20*18%,0)') // GST
    // LS for Unforeseen items is the reverse plug against the Grand Total,
    // not a forward sum — matching the real template exactly.
    expect(cellText(ws, 22, 2)).toBe('Add LS for Unforeseen items')
    expect(cellText(ws, 22, 12)).toBe('=L23-(L20+L21)')
    expect(cellText(ws, 23, 2)).toBe('Grand Total')
    expect(cellText(ws, 23, 12)).toBe('=CEILING(L20+L21,10000)')
  })

  it('writes the signature block with the Circle name', async () => {
    const ws = await loadSheet(await buildDetailedEstimateWorkbook(items, { circle: 'Nizampet Circle-58' }))
    let foundAe = false
    let foundCircle = false
    for (let r = 1; r <= ws.rowCount; r++) {
      if (cellText(ws, r, 2).trim() === 'Assistant Engineer') foundAe = true
      if (cellText(ws, r, 2) === 'Nizampet Circle-58') foundCircle = true
    }
    expect(foundAe).toBe(true)
    expect(foundCircle).toBe(true)
  })

  it('produces a valid workbook with a zero Sub Total when there are no items', async () => {
    const ws = await loadSheet(await buildDetailedEstimateWorkbook([]))
    expect(cellText(ws, 6, 2)).toBe('Sub Total ')
    expect(cellText(ws, 6, 12)).toBe('0')
  })
})

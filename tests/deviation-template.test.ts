import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import * as ExcelJS from 'exceljs'
import { fillDeviationTemplate } from '../core/deviationTemplate'
import type { DeviationItem } from '../core/deviationTemplate'

const TEMPLATE_PATH = resolve(__dirname, '../resources/deviation-template.xlsx')

function item(description: string, unit: string, quantity: string, rate: string): DeviationItem {
  return { description, unit, quantity, rate }
}

async function loadSheet(buffer: Buffer) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buffer as unknown as ArrayBuffer)
  return wb.worksheets[0]
}

function formulaAt(ws: ExcelJS.Worksheet, row: number, col: number): string | null {
  const v = ws.getCell(row, col).value as unknown
  if (v && typeof v === 'object' && 'formula' in v) return (v as { formula: string }).formula
  return null
}

describe('fillDeviationTemplate', () => {
  const templateBuffer = readFileSync(TEMPLATE_PATH)

  it('fills the header meta (Circle, Name of Work, Agency, Estimate Amount in Lakhs)', async () => {
    const out = await fillDeviationTemplate(
      templateBuffer,
      [item('Earth work excavation', 'Cum', '10', '100')],
      { circle: 'Gajularamaram Circle-57', nameOfWork: 'Test road work', agencyName: 'M/s Test Agency', estimateAmountLakhs: 12 }
    )
    const ws = await loadSheet(out)
    expect(ws.getCell(2, 1).value).toBe('Gajularamaram Circle-57')
    expect(ws.getCell(4, 3).value).toBe('Test road work')
    expect(ws.getCell(5, 3).value).toBe('M/s Test Agency')
    expect(ws.getCell(6, 3).value).toBe(12)
  })

  it('fills the estimate side of each item row, leaving the work-done side blank', async () => {
    const out = await fillDeviationTemplate(
      templateBuffer,
      [
        item('Earth work excavation', 'Cum', '10', '100'),
        item('PCC M20', 'Cum', '5', '6435.24')
      ],
      { estimateAmountLakhs: 1 }
    )
    const ws = await loadSheet(out)

    expect(ws.getCell(9, 1).value).toBe(1) // Sl No
    expect(ws.getCell(9, 2).value).toBe('Earth work excavation')
    expect(ws.getCell(9, 3).value).toBe('Cum') // Unit (estimate)
    expect(ws.getCell(9, 4).value).toBe(10) // Qty (estimate)
    expect(ws.getCell(9, 5).value).toBe(100) // Rate (estimate)
    expect(formulaAt(ws, 9, 6)).toBe('ROUND(D9*E9,0)') // Amount (estimate) — live formula

    expect(ws.getCell(9, 7).value).toBe('Cum') // Unit (work done) — copied
    expect(ws.getCell(9, 8).value).toBeFalsy() // Qty (work done) — blank
    expect(ws.getCell(9, 9).value).toBeFalsy() // Rate (work done) — blank
    expect(ws.getCell(9, 10).value).toBeFalsy() // Amount (work done) — blank

    expect(formulaAt(ws, 9, 11)).toBe('IF(J9>F9,J9-F9,0)') // Excess — kept live
    expect(formulaAt(ws, 9, 12)).toBe('IF(F9>J9,F9-J9,0)') // Less — kept live

    expect(ws.getCell(10, 1).value).toBe(2)
    expect(ws.getCell(10, 2).value).toBe('PCC M20')
  })

  it('re-targets the whole downstream cascade when the item count matches the template exactly (25, delta 0)', async () => {
    const items = Array.from({ length: 25 }, (_, i) => item(`Item ${i + 1}`, 'Cum', '10', '100'))
    const out = await fillDeviationTemplate(templateBuffer, items, { estimateAmountLakhs: 25 })
    const ws = await loadSheet(out)

    // Sub Total still at row 34 (25 items, same as the template's own).
    expect(formulaAt(ws, 34, 6)).toBe('SUM(F9:F33)')
    expect(formulaAt(ws, 34, 10)).toBe('SUM(J9:J33)')
    // Labour Cess (row 39) still references the Sub Total at row 34.
    expect(formulaAt(ws, 39, 6)).toBe('ROUND(F34*1%,0)')
    // GST (estimate side, row 47) is now a live formula off the Total-of-additions row (46).
    expect(formulaAt(ws, 47, 6)).toBe('ROUND(F46*18%,0)')
    // Final TOTAL (row 49) is the estimate amount in rupees.
    expect(ws.getCell(49, 6).value).toBe(2500000)
    // Seigniorage sub-table's item-specific QTY cells are cleared.
    expect(ws.getCell(60, 3).value).toBeFalsy()
    expect(ws.getCell(61, 3).value).toBeFalsy()
  })

  it('shifts every downstream formula correctly when there are FEWER items than the template (delta < 0)', async () => {
    const items = [item('Only item', 'Cum', '10', '100')] // 1 item, delta = 1 - 25 = -24
    const out = await fillDeviationTemplate(templateBuffer, items, { estimateAmountLakhs: 5 })
    const ws = await loadSheet(out)

    const expectedSubTotalRow = 9 + 1 // FIRST_ITEM_ROW + n
    expect(formulaAt(ws, expectedSubTotalRow, 6)).toBe(`SUM(F9:F9)`)
    expect(formulaAt(ws, expectedSubTotalRow, 10)).toBe(`SUM(J9:J9)`)

    const labourCessRow = expectedSubTotalRow + 5
    expect(formulaAt(ws, labourCessRow, 6)).toBe(`ROUND(F${expectedSubTotalRow}*1%,0)`)

    const gstRow = expectedSubTotalRow + 13
    const totalOfAdditionsRow = expectedSubTotalRow + 12
    expect(formulaAt(ws, gstRow, 6)).toBe(`ROUND(F${totalOfAdditionsRow}*18%,0)`)

    const finalTotalRow = expectedSubTotalRow + 15
    expect(ws.getCell(finalTotalRow, 6).value).toBe(500000)

    const seigniorageRow1 = expectedSubTotalRow + 26
    const seigniorageRow2 = expectedSubTotalRow + 27
    expect(ws.getCell(seigniorageRow1, 3).value).toBeFalsy()
    expect(ws.getCell(seigniorageRow2, 3).value).toBeFalsy()

    // The final signature block (last row of the template) still has its own text, just relocated.
    const lastRow = expectedSubTotalRow + 47
    const rowText = [1, 2, 3, 4, 5, 6, 7].map((c) => String(ws.getCell(lastRow, c).value ?? '')).join('')
    expect(rowText).toContain('AE')
  })

  it('shifts every downstream formula correctly when there are MORE items than the template (delta > 0)', async () => {
    const items = Array.from({ length: 30 }, (_, i) => item(`Item ${i + 1}`, 'Cum', '10', '100')) // 30 items, delta = +5
    const out = await fillDeviationTemplate(templateBuffer, items, { estimateAmountLakhs: 30 })
    const ws = await loadSheet(out)

    const expectedSubTotalRow = 9 + 30
    expect(formulaAt(ws, expectedSubTotalRow, 6)).toBe(`SUM(F9:F38)`)

    const labourCessRow = expectedSubTotalRow + 5
    expect(formulaAt(ws, labourCessRow, 6)).toBe(`ROUND(F${expectedSubTotalRow}*1%,0)`)

    const gstRow = expectedSubTotalRow + 13
    const totalOfAdditionsRow = expectedSubTotalRow + 12
    expect(formulaAt(ws, gstRow, 6)).toBe(`ROUND(F${totalOfAdditionsRow}*18%,0)`)

    const finalTotalRow = expectedSubTotalRow + 15
    expect(ws.getCell(finalTotalRow, 6).value).toBe(3000000)

    // The 30th (last) item row filled correctly.
    expect(ws.getCell(9 + 29, 2).value).toBe('Item 30')
    expect(formulaAt(ws, 9 + 29, 6)).toBe('ROUND(D38*E38,0)')
  })

  // The abstract labels (Add Labour Cess / QCC / NAC / … / GST) sit in three
  // cells each (C,D,E and G,H,I); a C:E / G:I merge hides the duplicates so the
  // label reads once. Resizing the item table used to drop those merges,
  // leaving the label repeated across three columns.
  it('keeps the abstract labels merged across C:E / G:I and centred at any item count', async () => {
    for (const n of [3, 25, 30]) {
      const items = Array.from({ length: n }, (_, i) => item(`Item ${i + 1}`, 'Cum', '10', '100'))
      const out = await fillDeviationTemplate(templateBuffer, items, { estimateAmountLakhs: 5 })
      const ws = await loadSheet(out)
      const merges = ws.model.merges as string[]

      const subTotalRow = 9 + n
      const labourCessRow = subTotalRow + 5 // estimate-side label span C:E
      const gstRow = subTotalRow + 13 // work-done-side label span G:I

      expect(ws.getCell(labourCessRow, 3).value).toBe('Add Labour Cess @ 1%')
      expect(merges).toContain(`C${labourCessRow}:E${labourCessRow}`)
      expect(ws.getCell(labourCessRow, 3).alignment?.horizontal).toBe('center')

      expect(ws.getCell(gstRow, 7).value).toBe('Add GST @ 18%')
      expect(merges).toContain(`G${gstRow}:I${gstRow}`)
      expect(ws.getCell(gstRow, 7).alignment?.horizontal).toBe('center')
    }
  })
})

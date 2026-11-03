import { describe, expect, it } from 'vitest'
import * as ExcelJS from 'exceljs'
import { excelToPrintHtml, inspectExcelForPrint } from '../core/excelPrintHtml'

async function book(build: (wb: ExcelJS.Workbook) => void): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()
  build(wb)
  return Buffer.from(await wb.xlsx.writeBuffer())
}

describe('excelToPrintHtml', () => {
  it('trims empty edges, keeps merges, and picks portrait for a narrow sheet', async () => {
    const xlsx = await book((wb) => {
      const ws = wb.addWorksheet('Narrow')
      ws.getColumn(2).width = 20
      ws.getColumn(3).width = 12
      ws.mergeCells('B2:C2')
      ws.getCell('B2').value = 'Title'
      ws.getCell('B3').value = 'Rate'
      ws.getCell('C3').value = 1234.5
      ws.getCell('C3').numFmt = '#,##0.00'
    })
    const sheets = await inspectExcelForPrint(xlsx)
    expect(sheets).toEqual([expect.objectContaining({ name: 'Narrow', rows: 2, cols: 2, orientation: 'portrait' })])
    const { html } = await excelToPrintHtml(xlsx)
    expect(html).toContain('colspan="2"')
    expect(html).toContain('1,234.50')
    expect(html).toContain('size: A4 portrait')
  })

  it('goes landscape for a wide sheet and honours the sheet selection', async () => {
    const xlsx = await book((wb) => {
      const wide = wb.addWorksheet('Wide')
      for (let c = 1; c <= 14; c++) {
        wide.getColumn(c).width = 14
        wide.getCell(1, c).value = `H${c}`
      }
      wb.addWorksheet('Other').getCell('A1').value = 'x'
    })
    const { html, sheets } = await excelToPrintHtml(xlsx, { sheets: ['Wide'] })
    expect(sheets).toBe(1)
    expect(html).toContain('size: A4 landscape')
    expect(html).not.toContain('>x<')
  })
})

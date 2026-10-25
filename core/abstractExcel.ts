import * as ExcelJS from 'exceljs'
import type { MonitoringFormatRow, MonitoringFormatSummary } from './monitoringFormat'
import { ABSTRACT_COL_UNITS, ABSTRACT_GROUP_KEYS } from './worksListPdf'

/**
 * The Monitoring Format Abstract as an .xlsx laid out like the PDF / online sheet: a merged title bar, the
 * "Name of the Circle / Date" line (values red), the three-row merged heading, item-type rows + Total, a blank row,
 * then the same heading over the ward/circle-wise rows. Legal landscape, one page wide; rows auto-fit their text.
 */
export async function buildAbstractWorkbookBuffer(summary: MonitoringFormatSummary, sheetTitle: string): Promise<Buffer> {
  const circleLevel = summary.wardPivotKind !== 'circle'
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Abstract', {
    pageSetup: { orientation: 'landscape', paperSize: 5, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 } }
  })
  ws.columns = ABSTRACT_COL_UNITS.map((u) => ({ width: u }))
  const thin = { style: 'thin' as const, color: { argb: 'FF000000' } }
  const box = { top: thin, left: thin, bottom: thin, right: thin }
  const red = { argb: 'FFFF0000' }
  const col = (c: number) => c + 1 // 0-based PDF column -> 1-based sheet column

  const put = (
    r: number,
    c0: number,
    c1: number,
    r1: number,
    text: string | number,
    o: { bold?: boolean; size?: number; align?: 'left' | 'center' | 'right'; color?: { argb: string }; border?: boolean } = {}
  ) => {
    if (c1 > c0 || r1 > r) ws.mergeCells(r, col(c0), r1, col(c1))
    const cell = ws.getCell(r, col(c0))
    cell.value = text
    cell.font = { name: 'Times New Roman', bold: o.bold, size: o.size ?? 11, color: o.color }
    cell.alignment = { wrapText: true, vertical: 'middle', horizontal: o.align ?? 'center' }
    if (o.border !== false) for (let rr = r; rr <= r1; rr++) for (let cc = c0; cc <= c1; cc++) ws.getCell(rr, col(cc)).border = box
  }

  let r = 1
  put(r, 0, 22, r, sheetTitle, { bold: true, size: 14 })
  ws.getRow(r).height = 30
  r++
  put(r, 1, 3, r, circleLevel ? 'Name of the Circle:-' : 'Name of the Zone:-', { bold: true, size: 12, border: false })
  put(r, 4, 7, r, summary.officeLabel, { bold: true, size: 12, color: red, border: false })
  put(r, 18, 19, r, 'Date:-', { bold: true, size: 12, border: false })
  put(r, 20, 22, r, summary.asOfDate ?? '', { bold: true, size: 12, color: red, border: false })
  for (let c = 0; c <= 22; c++) ws.getCell(r, col(c)).border = { top: thin, bottom: thin }
  ws.getRow(r).height = 22
  r++

  const table = (firstHeading: string, rows: MonitoringFormatRow[], totals: MonitoringFormatRow | undefined) => {
    const h = r
    put(h, 0, 0, h + 2, firstHeading, { bold: true })
    const groups: [number, number, string][] = [[1, 2, 'Total works'], [3, 4, 'Completed'], [15, 16, 'To be Started'], [17, 18, 'Tender Process'], [19, 20, 'Held Up'], [21, 22, 'Cancelled']]
    groups.forEach(([c0, c1, t]) => put(h, c0, c1, h + 1, t, { bold: true }))
    put(h, 5, 12, h, 'Progress', { bold: true })
    put(h, 13, 14, h, '')
    ;([[5, 6, 'upto 25%'], [7, 8, 'upto 50%'], [9, 10, 'upto 75%'], [11, 12, 'Above 75%'], [13, 14, 'TOTAL']] as [number, number, string][]).forEach(([c0, c1, t]) =>
      put(h + 1, c0, c1, h + 1, t, { bold: true })
    )
    for (let c = 1; c <= 22; c += 2) {
      put(h + 2, c, c, h + 2, 'No.', { bold: true })
      put(h + 2, c + 1, c + 1, h + 2, 'Amt in Rs Lakhs', { bold: true })
    }
    ws.getRow(h + 2).height = 48
    r += 3
    const body = [...rows.map((x) => ({ x, name: x.itemType, total: false })), ...(totals ? [{ x: totals, name: 'Total', total: true }] : [])]
    for (const { x, name, total } of body) {
      put(r, 0, 0, r, name, { bold: true, align: 'left' })
      ABSTRACT_GROUP_KEYS.forEach((k, g) => {
        ;[x[k].no, x[k].amt].forEach((v, j) => {
          put(r, 1 + g * 2 + j, 1 + g * 2 + j, r, Math.round((Number.isFinite(v) ? v : 0) * 100) / 100, { bold: total, align: 'right' })
        })
      })
      r++
    }
  }
  table('Item type', summary.rows, summary.totals)
  if (summary.wardRows && summary.wardRows.length > 0) {
    r++ // one clear blank row between the two tables
    table(summary.wardPivotKind === 'circle' ? 'Circle Name and number' : 'Ward Name and number', summary.wardRows, summary.wardTotals)
  }
  return Buffer.from(await wb.xlsx.writeBuffer())
}

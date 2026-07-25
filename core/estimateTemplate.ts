import * as ExcelJS from 'exceljs'
import type { EstimateWorkItem } from './estimateExtract'

/** Context filled into the header/signature blocks — all optional, since a photo-derived estimate may not resolve every field. */
export interface DetailedEstimateMeta {
  zone?: string
  circle?: string
  workName?: string
}

const LAST_COL = 12 // A..L, matching the department's own "Sl.No / Description / No's / (x) / L / B / D / Qty / Rate / Per / Amount" layout
const SNO_COL = 1
const DESC_COL = 2
const NOS_COL = 3
const L_COL = 6
const B_COL = 7
const D_COL = 8
const QTY_COL = 9
const RATE_COL = 10
const UNIT_COL = 11
const AMOUNT_COL = 12
// Column widths, indexed 1..12 — the merged Description block's total width
// (columns 2 through 8) is derived from these for the row-height estimate below.
const COL_WIDTHS = [6, 46, 6, 4, 4, 9, 9, 9, 10, 11, 8, 14]

// The department's own template (a real, in-use estimate shared directly
// for this) is Times New Roman 11pt throughout, bold for the letterhead/
// header/total rows — matched here exactly rather than this app's usual
// Verdana, since this is meant to read as the same document, not a
// reformatted approximation of it.
const FONT_NAME = 'Times New Roman'
const FONT_SIZE = 11

const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin' },
  bottom: { style: 'thin' },
  left: { style: 'thin' },
  right: { style: 'thin' }
}

function toNumber(s: string): number {
  const n = Number(String(s).replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : 0
}

function colLetter(col: number): string {
  let s = ''
  let n = col
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/**
 * Reassigns the whole `cell.style` object rather than setting `.font` /
 * `.alignment` / `.border` directly — ExcelJS shares one style object by
 * reference across cells that haven't been individually styled yet, so
 * mutating a sub-property in place can silently affect other cells still
 * pointing at that same shared object, not just the one cell this call is
 * meant to touch (the exact bug core/boqTemplate.ts's styleCell exists to
 * avoid — same fix, applied here too).
 */
function styleCell(
  cell: ExcelJS.Cell,
  style: { font?: Partial<ExcelJS.Font>; alignment?: Partial<ExcelJS.Alignment>; border?: Partial<ExcelJS.Borders> }
): void {
  cell.style = {
    ...cell.style,
    ...(style.font && { font: { ...cell.style.font, ...style.font } }),
    ...(style.alignment && { alignment: { ...cell.style.alignment, ...style.alignment } }),
    ...(style.border && { border: { ...cell.style.border, ...style.border } })
  }
}

function mergedText(ws: ExcelJS.Worksheet, row: number, text: string, bold = false): void {
  ws.mergeCells(row, 1, row, LAST_COL)
  const cell = ws.getCell(row, 1)
  cell.value = text
  styleCell(cell, { font: { name: FONT_NAME, size: FONT_SIZE, bold }, alignment: { horizontal: 'center', vertical: 'middle', wrapText: true } })
}

/** A totals-cascade row spanning the description columns (B:K), with its figure in the Amount column — matches the template's own Sub Total/LC/GST/Grand Total row shape. */
function chargeRow(ws: ExcelJS.Worksheet, row: number, label: string, value: number | { formula: string }, bold = false): void {
  ws.mergeCells(row, 2, row, LAST_COL - 1)
  const labelCell = ws.getCell(row, 2)
  labelCell.value = label
  styleCell(labelCell, { font: { name: FONT_NAME, size: FONT_SIZE, bold } })
  const amountCell = ws.getCell(row, AMOUNT_COL)
  amountCell.value = value
  styleCell(amountCell, { font: { name: FONT_NAME, size: FONT_SIZE, bold }, alignment: { horizontal: 'right', vertical: 'middle' } })
  amountCell.numFmt = '#,##0.00'
}

// Rough Excel-column-width-unit-to-character conversion for Times New Roman
// 11pt, used only to pick a row height tall enough to show a wrapped
// description in full — deliberately generous (fewer characters assumed per
// line than will usually actually fit) so the estimate errs toward a row
// that's a little taller than necessary rather than one that clips text,
// since this app doesn't render/measure the font itself to compute this
// exactly the way a spreadsheet application does.
const CHARS_PER_WIDTH_UNIT = 1.6
const LINE_HEIGHT_POINTS = 14

/** How tall (in points) a row needs to be for `text` to fully display, wrapped, within a merged block `mergedWidthUnits` wide — computed explicitly rather than left to the consuming application's own auto-fit, since not every spreadsheet application reliably auto-sizes a merged, wrapped cell's row height (Google Sheets in particular does not). */
function wrappedRowHeight(text: string, mergedWidthUnits: number): number {
  const charsPerLine = Math.max(10, Math.floor(mergedWidthUnits * CHARS_PER_WIDTH_UNIT))
  const lines = Math.max(1, Math.ceil(text.length / charsPerLine))
  return lines * LINE_HEIGHT_POINTS + 4
}

/**
 * Build a "Detailed and Abstract Estimate" workbook matching the
 * department's own real template exactly (letterhead, title, name of work,
 * estimate amount, an Sl.No/Description/Qty/Rate/Per/Amount item table, and
 * the standard surcharge cascade below it) — not a bare Sl No/Description/
 * Unit/Quantity/Rate/Amount table.
 *
 * The totals cascade mirrors the real template's own formulas exactly:
 * Sub Total (sum of item amounts, i.e. the ECV) → LC 1% → NAC 0.1% →
 * TPQC 0.35% → **Seigniorage charges = 2% of the Sub Total** → DMFT 30%/
 * SMET 2%/Permit Fee 80% (all of Seigniorage) → a second Sub Total → GST
 * 18% → **Grand Total, a manually-entered administratively-sanctioned
 * figure** (seeded here to the next round ₹10,000 above the computed
 * total, editable) → **"LS for Unforeseen items", the reverse plug that
 * makes Sub Total 2 + GST + LS reconcile exactly to that Grand Total** —
 * matching how the real template computes it (`=GrandTotal-(SubTotal2+GST)`),
 * not the other way around. The header's own "Estimate Amount ... lakhs"
 * is a live formula off the Grand Total too, same as the original.
 */
export async function buildDetailedEstimateWorkbook(
  items: EstimateWorkItem[],
  meta: DetailedEstimateMeta = {}
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  const ws = workbook.addWorksheet('Estimate')
  ws.columns = COL_WIDTHS.map((width) => ({ width }))
  const descMergedWidth = COL_WIDTHS.slice(DESC_COL - 1, QTY_COL).reduce((a, b) => a + b, 0)

  let r = 1
  const letterhead = [meta.circle, meta.zone ? `${meta.zone} ZONE` : undefined].filter(Boolean).join(' :: ')
  mergedText(ws, r++, letterhead || 'MUNICIPAL CORPORATION', true)
  mergedText(ws, r++, 'Detailed and Abstract - ESTIMATE', true)
  mergedText(ws, r++, `Name of the work: ${meta.workName ?? ''}`, true)

  // Estimate Amount is a live readout of the Grand Total below, not a
  // separately-entered figure — filled in once the Grand Total row's own
  // row number is known (see estimateAmountCell below).
  ws.mergeCells(r, 6, r, 10)
  const amountLabelCell = ws.getCell(r, 6)
  amountLabelCell.value = 'Estimate Amount Rs. '
  styleCell(amountLabelCell, { font: { name: FONT_NAME, size: FONT_SIZE, bold: true }, alignment: { horizontal: 'right' } })
  const estimateAmountRow = r
  const lakhsCell = ws.getCell(r, 12)
  lakhsCell.value = 'lakhs'
  styleCell(lakhsCell, { font: { name: FONT_NAME, size: FONT_SIZE, bold: true } })
  r++

  const headerRow = r
  ws.mergeCells(headerRow, 3, headerRow, 5)
  const headers = ['Sl. No.', 'Description of work', "No's", '', '', 'L', 'B', 'D', 'Qty', 'Rate', 'Per', 'Amount']
  headers.forEach((h, i) => {
    if (i + 1 === 4 || i + 1 === 5) return // part of the No's merge above
    const cell = ws.getCell(headerRow, i + 1)
    cell.value = h
    styleCell(cell, {
      font: { name: FONT_NAME, size: FONT_SIZE, bold: true },
      alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
      border: THIN_BORDER
    })
  })
  r++

  // Each item spans 3 rows, matching the real template exactly: a lead row
  // (Sl.No + Description, spanning columns B through I since nothing else
  // sits on that row), a dimension row (No's/L/B/D, filled in when the
  // source estimate shows that breakdown, left blank otherwise), and a
  // summary row (Qty/Rate/Per/Amount, immediately below the description,
  // not sharing its row).
  const ROWS_PER_ITEM = 3
  const firstItemRow = r
  items.forEach((item, i) => {
    const leadRow = firstItemRow + i * ROWS_PER_ITEM
    const dimRow = leadRow + 1
    const summaryRow = leadRow + 2

    ws.mergeCells(leadRow, DESC_COL, leadRow, QTY_COL)
    ws.getCell(leadRow, SNO_COL).value = i + 1
    const descCell = ws.getCell(leadRow, DESC_COL)
    descCell.value = item.description
    for (let c = 1; c <= LAST_COL; c++) {
      const alignment: Partial<ExcelJS.Alignment> =
        c === DESC_COL ? { wrapText: true, vertical: 'top', horizontal: 'left' } : { horizontal: 'center', vertical: 'top' }
      styleCell(ws.getCell(leadRow, c), { font: { name: FONT_NAME, size: FONT_SIZE }, alignment, border: THIN_BORDER })
    }
    // Set explicitly rather than left to auto-fit — a merged, wrapped cell's
    // row height isn't reliably auto-computed by every spreadsheet
    // application (see wrappedRowHeight's doc comment).
    ws.getRow(leadRow).height = wrappedRowHeight(item.description, descMergedWidth)

    for (let c = 1; c <= LAST_COL; c++) {
      styleCell(ws.getCell(dimRow, c), {
        font: { name: FONT_NAME, size: FONT_SIZE },
        alignment: { horizontal: 'center', vertical: 'middle' },
        border: THIN_BORDER
      })
    }
    if (item.nos) ws.getCell(dimRow, NOS_COL).value = toNumber(item.nos)
    if (item.l) ws.getCell(dimRow, L_COL).value = toNumber(item.l)
    if (item.b) ws.getCell(dimRow, B_COL).value = toNumber(item.b)
    if (item.d) ws.getCell(dimRow, D_COL).value = toNumber(item.d)

    const qtyCell = ws.getCell(summaryRow, QTY_COL)
    qtyCell.value = toNumber(item.quantity)
    const rateCell = ws.getCell(summaryRow, RATE_COL)
    rateCell.value = toNumber(item.rate)
    const unitCell = ws.getCell(summaryRow, UNIT_COL)
    unitCell.value = item.unit
    const amountCell = ws.getCell(summaryRow, AMOUNT_COL)
    amountCell.value = { formula: `ROUND(${colLetter(QTY_COL)}${summaryRow}*${colLetter(RATE_COL)}${summaryRow},0)` }
    for (let c = 1; c <= LAST_COL; c++) {
      styleCell(ws.getCell(summaryRow, c), {
        font: { name: FONT_NAME, size: FONT_SIZE, bold: c === QTY_COL },
        alignment: { horizontal: c === QTY_COL || c === RATE_COL || c === AMOUNT_COL ? 'right' : 'center', vertical: 'middle' },
        border: THIN_BORDER
      })
    }
    amountCell.numFmt = '#,##0.00'
    rateCell.numFmt = '#,##0.00'
  })
  const lastItemRow = firstItemRow + items.length * ROWS_PER_ITEM - 1

  r = lastItemRow + 1
  const subTotal1Row = r
  chargeRow(
    ws,
    r,
    'Sub Total ',
    items.length > 0 ? { formula: `SUM(${colLetter(AMOUNT_COL)}${firstItemRow}:${colLetter(AMOUNT_COL)}${lastItemRow})` } : 0,
    true
  )
  r++

  const lcRow = r
  chargeRow(ws, r, 'LC 1.00%', { formula: `ROUND(${colLetter(AMOUNT_COL)}${subTotal1Row}*1%,0)` })
  r++
  chargeRow(ws, r, 'Add 0.10% N.A.C', { formula: `ROUND(${colLetter(AMOUNT_COL)}${subTotal1Row}*0.1%,0)` })
  r++
  chargeRow(ws, r, 'TPQC 0.35%', { formula: `ROUND(${colLetter(AMOUNT_COL)}${subTotal1Row}*0.35%,0)` })
  r++
  // 2% of the Sub Total (the item amounts' total, i.e. the ECV) — matching
  // the real template's own formula exactly, not left for manual entry.
  const seigniorageRow = r
  chargeRow(ws, r, 'Seigniorage charges ', { formula: `ROUND(${colLetter(AMOUNT_COL)}${subTotal1Row}*2%,0)` })
  r++
  chargeRow(ws, r, 'DMFT 30%', { formula: `ROUND(30%*${colLetter(AMOUNT_COL)}${seigniorageRow},0)` })
  r++
  chargeRow(ws, r, 'SMET 2%', { formula: `ROUND(${colLetter(AMOUNT_COL)}${seigniorageRow}*2%,0)` })
  r++
  const permitFeeRow = r
  chargeRow(ws, r, 'Permit fee  0.8 on Seignorage', { formula: `ROUND(${colLetter(AMOUNT_COL)}${seigniorageRow}*0.8,0)` })
  r++

  const subTotal2Row = r
  chargeRow(ws, r, 'Sub Total ', { formula: `SUM(${colLetter(AMOUNT_COL)}${subTotal1Row}:${colLetter(AMOUNT_COL)}${permitFeeRow})` }, true)
  r++

  const gstRow = r
  chargeRow(ws, r, 'Add 18.00% GST', { formula: `ROUND(${colLetter(AMOUNT_COL)}${subTotal2Row}*18%,0)` })
  r++

  // Grand Total is entered here as the administratively-sanctioned figure —
  // seeded to the next round ₹10,000 above the computed total as a
  // reasonable starting point, but this cell is meant to be typed over with
  // whatever figure is actually sanctioned (matching the real template,
  // where it's a plain typed-in number, not a formula).
  const grandTotalRow = r + 1
  chargeRow(ws, r, 'Add LS for Unforeseen items', {
    formula: `${colLetter(AMOUNT_COL)}${grandTotalRow}-(${colLetter(AMOUNT_COL)}${subTotal2Row}+${colLetter(AMOUNT_COL)}${gstRow})`
  })
  r++

  chargeRow(
    ws,
    r,
    'Grand Total',
    { formula: `CEILING(${colLetter(AMOUNT_COL)}${subTotal2Row}+${colLetter(AMOUNT_COL)}${gstRow},10000)` },
    true
  )
  r += 3

  const estimateAmountCell = ws.getCell(estimateAmountRow, 11)
  estimateAmountCell.value = { formula: `ROUND(${colLetter(AMOUNT_COL)}${grandTotalRow}/100000,2)` }
  styleCell(estimateAmountCell, { font: { name: FONT_NAME, size: FONT_SIZE, bold: true } })
  estimateAmountCell.numFmt = '0.00'

  const sigRow = r
  const aeCell = ws.getCell(sigRow, 2)
  aeCell.value = 'Assistant Engineer '
  const deeCell = ws.getCell(sigRow, 8)
  deeCell.value = 'Deputy Executive Engineer'
  styleCell(aeCell, { font: { name: FONT_NAME, size: FONT_SIZE, bold: true } })
  styleCell(deeCell, { font: { name: FONT_NAME, size: FONT_SIZE, bold: true } })
  if (meta.circle) {
    const aeCircleCell = ws.getCell(sigRow + 1, 2)
    aeCircleCell.value = meta.circle
    const deeCircleCell = ws.getCell(sigRow + 1, 8)
    deeCircleCell.value = meta.circle
    styleCell(aeCircleCell, { font: { name: FONT_NAME, size: FONT_SIZE } })
    styleCell(deeCircleCell, { font: { name: FONT_NAME, size: FONT_SIZE } })
  }

  const out = await workbook.xlsx.writeBuffer()
  return Buffer.from(out)
}

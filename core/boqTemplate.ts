import * as ExcelJS from 'exceljs'
import type { ExcelTable } from './types'
import { BOQ_HEADERS } from './boqHeaders'
import { stripDataValidations, trimToContent } from './templateWorkbook'

// Fixed column layout of the bundled BOQ template.
const QTY_COL = 1
const DESC_COL = 2
const WORK_TYPE_COL = 3
const SHORT_DESC_COL = 4
const APSS_COL = 5
const RATE_COL = 6
const UOM_COL = 7
const AMOUNT_COL = 8

export interface BoqRowData {
  quantity: string
  description: string
  workType: string
  shortDescription: string
  apss: string
  rate: string
  uom: string
}

/** Pull the BOQ row data back out of the plain table built by buildBoqFromEstimate. */
export function rowsToBoqData(table: ExcelTable): BoqRowData[] {
  const [qtyH, descH, workTypeH, shortDescH, apssH, rateH, uomH] = BOQ_HEADERS
  return table.rows.map((row) => ({
    quantity: row[qtyH] ?? '',
    description: row[descH] ?? '',
    workType: row[workTypeH] ?? '',
    shortDescription: row[shortDescH] ?? '',
    apss: row[apssH] ?? '',
    rate: row[rateH] ?? '',
    uom: row[uomH] ?? ''
  }))
}

function cellText(v: ExcelJS.CellValue): string {
  if (v == null) return ''
  if (typeof v === 'object') {
    const rich = v as { richText?: { text: string }[] }
    if (rich.richText) return rich.richText.map((r) => r.text).join('')
    const formulaVal = v as { result?: unknown }
    if ('result' in (v as object)) return String(formulaVal.result ?? '')
  }
  return String(v)
}

function numOrStr(s: string | undefined): number | string {
  if (!s) return ''
  const n = Number(String(s).replace(/,/g, '').trim())
  return Number.isFinite(n) && String(s).trim() !== '' ? n : s
}

/**
 * Force Verdana 12pt, center/middle alignment — keeping every other
 * font/alignment attribute (bold, wrapText, …) as the template set it.
 *
 * Reassigns the whole `cell.style` object rather than `cell.font =` /
 * `cell.alignment =` directly: ExcelJS shares one style object by reference
 * across every cell that used the same style index in the loaded file (a
 * memory optimisation), so mutating a sub-property in place — which is what
 * the `.font`/`.alignment` setters do — silently corrupts every other cell
 * sharing that style, including unrelated rows far below the ones this
 * function is meant to touch. Assigning a fresh `.style` object breaks that
 * shared reference so only this cell changes.
 */
function styleCell(cell: ExcelJS.Cell): void {
  cell.style = {
    ...cell.style,
    font: { ...cell.style.font, name: 'Verdana', size: 12 },
    alignment: { ...cell.style.alignment, horizontal: 'center', vertical: 'middle' }
  }
}

/**
 * Let Excel auto-fit this row's height to its (now Verdana-12) content
 * instead of keeping the template's fixed height (tuned for the old size-10
 * font, which would clip wrapped text). ExcelJS only writes the `ht` /
 * `customHeight="1"` attributes — the ones that pin a row's height and stop
 * Excel from recalculating it — when `row.height` is set, so clearing it
 * omits both and Excel computes the height itself when the file is opened.
 */
function autoFitRowHeight(ws: ExcelJS.Worksheet, r: number): void {
  // ExcelJS's type declares `height: number` (non-optional), but the writer
  // only emits `ht`/`customHeight` when it's truthy — `undefined` at runtime
  // is what actually triggers auto-fit, the types just don't reflect that.
  ws.getRow(r).height = undefined as unknown as number
}

/**
 * Fill the bundled BOQ template in place — preserving its fill colours and
 * column widths — instead of rebuilding the document from scratch. Item rows
 * are inserted or removed to match the estimate's row count, using the
 * template's own item row as the formula reference, and the Amount column
 * stays a live formula (rate × qty, and a SUM total) rather than a static
 * number, matching the original template's design. Every header, item, and
 * total cell is forced to Verdana 12pt, center/middle aligned, and those
 * rows' heights are left to Excel's auto-fit rather than the template's
 * fixed heights (which were sized for the smaller original font).
 */
export async function fillBoqTemplate(
  templateBuffer: Buffer,
  rows: BoqRowData[],
  workName?: string
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(templateBuffer as unknown as ArrayBuffer)
  const ws = workbook.worksheets[0]
  if (!ws) throw new Error('BOQ template has no sheet.')
  stripDataValidations(ws)

  let headerRow = -1
  for (let r = 1; r <= Math.min(ws.rowCount, 10); r++) {
    if (cellText(ws.getCell(r, QTY_COL).value).trim().toLowerCase().startsWith('estimate quantity')) {
      headerRow = r
      break
    }
  }
  if (headerRow === -1) {
    throw new Error('Could not find the header row ("Estimate Quantity...") in the BOQ template.')
  }
  for (let c = QTY_COL; c <= AMOUNT_COL; c++) styleCell(ws.getCell(headerRow, c))
  autoFitRowHeight(ws, headerRow)

  const firstItemRow = headerRow + 1
  let templateItemCount = 0
  while (cellText(ws.getCell(firstItemRow + templateItemCount, DESC_COL).value).trim() !== '') {
    templateItemCount += 1
  }
  if (templateItemCount === 0) {
    throw new Error('BOQ template has no item row to use as a formatting reference.')
  }

  const n = rows.length
  if (n > templateItemCount) {
    ws.duplicateRow(firstItemRow + templateItemCount - 1, n - templateItemCount, true)
  } else if (n < templateItemCount) {
    ws.spliceRows(firstItemRow + n, templateItemCount - n)
  }

  for (let i = 0; i < n; i++) {
    const r = firstItemRow + i
    const it = rows[i]
    ws.getCell(r, QTY_COL).value = numOrStr(it.quantity)
    ws.getCell(r, DESC_COL).value = it.description
    ws.getCell(r, WORK_TYPE_COL).value = it.workType
    ws.getCell(r, SHORT_DESC_COL).value = it.shortDescription
    ws.getCell(r, APSS_COL).value = it.apss || 'NA'
    ws.getCell(r, RATE_COL).value = numOrStr(it.rate)
    ws.getCell(r, UOM_COL).value = it.uom
    ws.getCell(r, AMOUNT_COL).value = { formula: `ROUND(F${r}*A${r},0)` }
    for (let c = QTY_COL; c <= AMOUNT_COL; c++) styleCell(ws.getCell(r, c))
    autoFitRowHeight(ws, r)
  }

  const totalRow = firstItemRow + n
  ws.getCell(totalRow, AMOUNT_COL).value =
    n > 0 ? { formula: `SUM(H${firstItemRow}:H${totalRow - 1})` } : 0
  styleCell(ws.getCell(totalRow, AMOUNT_COL))
  autoFitRowHeight(ws, totalRow)

  // The estimate's own work name, extracted from its title block — written
  // two rows below the total (one blank row as a gap) as a plain reference,
  // not part of the item table itself.
  let lastContentRow = totalRow
  if (workName?.trim()) {
    lastContentRow = totalRow + 2
    const cell = ws.getCell(lastContentRow, DESC_COL)
    cell.value = `Name of Work: ${workName.trim()}`
    cell.style = { ...cell.style, font: { ...cell.style.font, name: 'Verdana', size: 12, bold: true } }
  }
  trimToContent(ws, lastContentRow, AMOUNT_COL)

  const out = await workbook.xlsx.writeBuffer()
  return Buffer.from(out)
}

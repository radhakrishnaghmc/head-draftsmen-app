import * as ExcelJS from 'exceljs'

// Fixed layout of the bundled Deviation Statement template
// (resources/deviation-template.xlsx) — a real, human-authored government
// financial document, not a generic table. Row numbers below are the
// template's own fixed structure (1-indexed, ExcelJS convention) and never
// change; only the item table's row COUNT varies per upload.
const CIRCLE_ROW = 2
const NAME_OF_WORK_ROW = 4
const NAME_OF_WORK_COL = 3 // C4:M4 merged
const AGENCY_ROW = 5
const AGENCY_COL = 3 // C5:M5 merged
const ESTIMATE_AMOUNT_ROW = 6
const ESTIMATE_AMOUNT_COL = 3 // C6:D6 merged — plain number, in Lakhs

const FIRST_ITEM_ROW = 9
const TEMPLATE_ITEM_COUNT = 25 // template's own sample row count (rows 9-33)

const SL_NO_COL = 1
const DESC_COL = 2
const UNIT_EST_COL = 3
const QTY_EST_COL = 4
const RATE_EST_COL = 5
const AMOUNT_EST_COL = 6
const UNIT_WD_COL = 7
const QTY_WD_COL = 8
const RATE_WD_COL = 9
const AMOUNT_WD_COL = 10
const EXCESS_COL = 11
const LESS_COL = 12

// Everything from the Sub Total row down is a cascade of formulas (Labour
// Cess, QCC, NAC, Seigniorage, DMFT, SMET, Permit Fee, GST, tender
// percentage, the final contract/work-done/difference summary, and a
// Seigniorage sub-table) that all reference each other by fixed row
// offsets from the Sub Total row — none of it reference the item rows
// directly except the two exceptions called out below. Offsets are
// (original row) - (original Sub Total row, 34).
const SUB_TOTAL_OFFSET = 0
// The Seigniorage sub-table's two "QTY" cells (C60/C61 in the original)
// reference *specific* item rows (a human classified which items are
// "pcc m20"/"Wet Mix" for royalty purposes) — that can't generalize to a
// new estimate's different items, so these are cleared instead of shifted.
const SEIGNIORAGE_PCC_QTY_OFFSET = 26
const SEIGNIORAGE_WETMIX_QTY_OFFSET = 27
// "Add GST @ 18%" (estimate side) and the final "TOTAL" (estimate side) are
// the template's own static numbers (not formulas) from its sample project
// — TOTAL is the work's own sanctioned estimate amount (matches the header
// field), and GST is recomputed as a live formula to match its work-done
// counterpart's own convention.
const GST_ESTIMATE_SIDE_OFFSET = 13
const GST_ESTIMATE_SIDE_COL = 6
const TOTAL_OF_ADDITIONS_OFFSET = 12
const FINAL_TOTAL_OFFSET = 15
const FINAL_TOTAL_COL = 6
// The template's last row with any content (the signature block).
const LAST_ROW_OFFSET = 47

export interface DeviationItem {
  description: string
  unit: string
  quantity: string
  rate: string
}

export interface DeviationMeta {
  circle?: string
  nameOfWork?: string
  agencyName?: string
  /** In Lakhs — matches how every other amount field in this app is entered. */
  estimateAmountLakhs: number
}

function cellText(v: ExcelJS.CellValue): string {
  if (v == null) return ''
  if (typeof v === 'object') {
    const rich = v as { richText?: { text: string }[] }
    if (rich.richText) return rich.richText.map((r) => r.text).join('')
  }
  return String(v)
}

function numOrStr(s: string | undefined): number | string {
  if (!s) return ''
  const n = Number(String(s).replace(/,/g, '').trim())
  return Number.isFinite(n) && String(s).trim() !== '' ? n : s
}

// -- Formula introspection: ExcelJS represents a cell's formula either
// directly ({formula: '...'}) or, for cells Excel optimized as part of a
// "shared formula" group, as a pointer to a master cell ({sharedFormula:
// 'I10'}) whose formula applies here too, adjusted by the same relative
// offset. Resolving through that pointer gives the *actual* formula any
// cell computes, regardless of which representation the file happened to
// use — necessary since the template mixes both.
const CELL_REF_RE = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g

function colLetterToNum(letters: string): number {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

function colNumToLetter(num: number): string {
  let s = ''
  let n = num
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** Shift every cell reference in a formula by a fixed row/column delta (used to re-derive a shared-formula cell's actual formula from its master). */
function shiftFormulaBy(formula: string, rowDelta: number, colDelta: number): string {
  return formula.replace(CELL_REF_RE, (_m, colAbs, col, rowAbs, rowStr) => {
    const newCol = colAbs ? col : colNumToLetter(colLetterToNum(col) + colDelta)
    const newRow = rowAbs ? rowStr : String(parseInt(rowStr, 10) + rowDelta)
    return `${colAbs}${newCol}${rowAbs}${newRow}`
  })
}

/** Shift only the cell references whose row is >= `threshold`, by `delta` — used to re-target the downstream cascade's cross-references after the item table's row count changes. */
function shiftFormulaAboveThreshold(formula: string, threshold: number, delta: number): string {
  return formula.replace(CELL_REF_RE, (m, colAbs, col, rowAbs, rowStr) => {
    const row = parseInt(rowStr, 10)
    if (row < threshold) return m
    return `${colAbs}${col}${rowAbs}${row + delta}`
  })
}

function effectiveFormula(ws: ExcelJS.Worksheet, row: number, col: number, depth = 0): string | null {
  if (depth > 5) return null
  const v = ws.getCell(row, col).value as unknown
  if (v && typeof v === 'object') {
    if ('formula' in v && (v as { formula?: string }).formula) return (v as { formula: string }).formula
    if ('sharedFormula' in v && (v as { sharedFormula?: string }).sharedFormula) {
      const masterAddr = (v as { sharedFormula: string }).sharedFormula
      const m = /^([A-Z]+)(\d+)$/.exec(masterAddr)
      if (!m) return null
      const masterCol = colLetterToNum(m[1])
      const masterRow = parseInt(m[2], 10)
      const masterFormula = effectiveFormula(ws, masterRow, masterCol, depth + 1)
      if (!masterFormula) return null
      return shiftFormulaBy(masterFormula, row - masterRow, col - masterCol)
    }
  }
  return null
}

/**
 * Fill the bundled Deviation Statement template: header meta (Circle, Name
 * of the work, Name of the Agency, Estimate Amount), and the item table's
 * "As per sanctioned estimate" side (Description/Unit/Qty/Rate copied from
 * the uploaded estimate, Amount as a live `=Qty*Rate` formula). The "As per
 * work done" side is left blank (that data doesn't exist until the work is
 * actually measured) — but every formula below it (Excess/Less, Sub Total,
 * Labour Cess, GST, Seigniorage, the final contract summary, ...) is kept
 * live, re-targeted to the new item count, so it all computes correctly
 * once work-done figures are filled in later.
 */
export async function fillDeviationTemplate(
  templateBuffer: Buffer,
  items: DeviationItem[],
  meta: DeviationMeta
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(templateBuffer as unknown as ArrayBuffer)
  const ws = workbook.worksheets[0]
  if (!ws) throw new Error('Deviation template has no sheet.')

  // Header meta.
  if (meta.circle) ws.getCell(CIRCLE_ROW, 1).value = meta.circle
  if (meta.nameOfWork) ws.getCell(NAME_OF_WORK_ROW, NAME_OF_WORK_COL).value = meta.nameOfWork
  if (meta.agencyName) ws.getCell(AGENCY_ROW, AGENCY_COL).value = meta.agencyName
  ws.getCell(ESTIMATE_AMOUNT_ROW, ESTIMATE_AMOUNT_COL).value = meta.estimateAmountLakhs

  const n = items.length
  const oldSubTotalRow = FIRST_ITEM_ROW + TEMPLATE_ITEM_COUNT
  const newSubTotalRow = FIRST_ITEM_ROW + n
  const delta = n - TEMPLATE_ITEM_COUNT
  const lastRow = oldSubTotalRow + LAST_ROW_OFFSET

  // Capture the downstream cascade's *actual* formulas (resolving shared-
  // formula pointers) before the item table's row count changes anything —
  // done first so every reference below is still relative to the
  // template's original, known-good layout.
  const downstream: { row: number; col: number; formula: string }[] = []
  for (let r = oldSubTotalRow; r <= lastRow; r++) {
    for (let c = 1; c <= 14; c++) {
      const f = effectiveFormula(ws, r, c)
      if (f) downstream.push({ row: r, col: c, formula: f })
    }
  }

  // Resize the item table to match the uploaded estimate, using the
  // template's own first item row as the formatting reference — same
  // approach as core/boqTemplate.ts.
  if (n > TEMPLATE_ITEM_COUNT) {
    ws.duplicateRow(FIRST_ITEM_ROW + TEMPLATE_ITEM_COUNT - 1, n - TEMPLATE_ITEM_COUNT, true)
  } else if (n < TEMPLATE_ITEM_COUNT) {
    ws.spliceRows(FIRST_ITEM_ROW + n, TEMPLATE_ITEM_COUNT - n)
  }

  for (let i = 0; i < n; i++) {
    const r = FIRST_ITEM_ROW + i
    const it = items[i]
    ws.getCell(r, SL_NO_COL).value = i + 1
    ws.getCell(r, DESC_COL).value = it.description
    ws.getCell(r, UNIT_EST_COL).value = it.unit
    ws.getCell(r, QTY_EST_COL).value = numOrStr(it.quantity)
    ws.getCell(r, RATE_EST_COL).value = numOrStr(it.rate)
    ws.getCell(r, AMOUNT_EST_COL).value = { formula: `ROUND(D${r}*E${r},0)` }
    ws.getCell(r, UNIT_WD_COL).value = it.unit
    ws.getCell(r, QTY_WD_COL).value = null
    ws.getCell(r, RATE_WD_COL).value = null
    ws.getCell(r, AMOUNT_WD_COL).value = null
    ws.getCell(r, EXCESS_COL).value = { formula: `IF(J${r}>F${r},J${r}-F${r},0)` }
    ws.getCell(r, LESS_COL).value = { formula: `IF(F${r}>J${r},F${r}-J${r},0)` }
  }

  // Re-target the downstream cascade to the new row positions.
  for (const { row, col, formula } of downstream) {
    const shifted = shiftFormulaAboveThreshold(formula, oldSubTotalRow, delta)
    ws.getCell(row + delta, col).value = { formula: shifted }
  }

  // Sub Total's own SUM ranges: the lower bound (the item table's first
  // row) never moves, but the upper bound must follow the new item count —
  // the generic shift above leaves it alone since it's below the old
  // Sub Total row's own threshold.
  const lastItemRow = FIRST_ITEM_ROW + n - 1
  ws.getCell(newSubTotalRow, AMOUNT_EST_COL).value = { formula: `SUM(F${FIRST_ITEM_ROW}:F${lastItemRow})` }
  ws.getCell(newSubTotalRow, AMOUNT_WD_COL).value = { formula: `SUM(J${FIRST_ITEM_ROW}:J${lastItemRow})` }
  ws.getCell(newSubTotalRow, EXCESS_COL).value = { formula: `SUM(K${FIRST_ITEM_ROW}:K${lastItemRow})` }
  ws.getCell(newSubTotalRow, LESS_COL).value = { formula: `SUM(L${FIRST_ITEM_ROW}:L${lastItemRow})` }

  // The Seigniorage sub-table's two "QTY" inputs reference specific item
  // rows from the template's own sample project — not generalizable to a
  // different estimate's items, so they're cleared rather than mis-shifted.
  ws.getCell(newSubTotalRow + SEIGNIORAGE_PCC_QTY_OFFSET, 3).value = null
  ws.getCell(newSubTotalRow + SEIGNIORAGE_WETMIX_QTY_OFFSET, 3).value = null

  // The final TOTAL (estimate side) is this work's own sanctioned estimate
  // amount — matches the header field, in rupees.
  ws.getCell(newSubTotalRow + FINAL_TOTAL_OFFSET, FINAL_TOTAL_COL).value = Math.round(meta.estimateAmountLakhs * 100000)
  // "Add GST @ 18%" (estimate side) was a static leftover number from the
  // template's own sample — recomputed as a live formula, matching its
  // work-done counterpart's own convention.
  const totalOfAdditionsRow = newSubTotalRow + TOTAL_OF_ADDITIONS_OFFSET
  ws.getCell(newSubTotalRow + GST_ESTIMATE_SIDE_OFFSET, GST_ESTIMATE_SIDE_COL).value = {
    formula: `ROUND(F${totalOfAdditionsRow}*18%,0)`
  }

  const out = await workbook.xlsx.writeBuffer()
  return Buffer.from(out)
}

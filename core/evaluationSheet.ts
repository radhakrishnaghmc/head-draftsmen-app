import * as ExcelJS from 'exceljs'

// Builds the "Evaluation Sheet" — the Bid Capacity evaluation the office runs on
// every participating bidder. One column per bidder, eleven fixed eligibility
// criteria, then the bid-capacity block: A (turnover), N (period, in years), B
// (existing commitments), 3AN−B (bid capacity), ECV and Remarks. The header,
// bidder columns and ECV are filled from the uploaded "View Bidders" sheet; the
// per-bidder data (turnover / N / B / eligibility / Remarks) is left blank for
// the office to key in from each bidder's documents, and the bid-capacity cell
// (3AN−B) plus the surplus are live formulas that resolve as those blanks fill.

/** The eleven standard tender-condition rows (column B, Sl. No 1–11). */
export const EVALUATION_CRITERIA: readonly string[] = [
  'Contractor registration (Class-V) issued by GHMC authorities only & Other Departments Class III & above as per G.O.M.S.No. 94 I & CAD, Dated: 01-07-2003 (Upload the original copy for clarity)',
  'GST Registration & Pan Card (Upload the original copy for clarity)',
  'The Contractor Should enclose Current Address & Mobile No.',
  'Value of existing commitments and ongoing works to be completed during the period of completion of works for which tenders are invited (if NIL certificate to be enclosed) and the same ongoing works will be verified with works projects',
  'The administrative sanction amount above 5.00 lakhs will be evaluated based on 3AN-B',
  'Maximum value of Civil Engineering works executed in any one year during the last 5 years updated price levels taking into account the completed as well as works in progress this value should be certified by C.A only will be accepted. CA certificate to be produced with only turnover of civil engineering works otherwise not accepted. (Upload the original copy for clarity)',
  'Income Tax Returns Certificate for the latest financial year (Upload the original copy for clarity)',
  'Caste Certificate to be submitted by the individuals other than agencies for reservation works (Caste certificate issued by the Telangana Government) (Upload the original copy for clarity)',
  'For reservation work only one work should be allotted for one agency. No another work will be allotted to the already allotted agency.',
  'Self-declaration of genuinity of documents to be uploaded by the agency',
  'The bidder shall submit litigation history.'
]

export interface EvaluationSheetInput {
  /** Full NIT line for row 2, e.g. "NIT No: 16/DB/EE/…/CMC/2026-27,Dt:25.07.2026,Tender ID:722231". */
  nitLine: string
  /** Full work line for row 3, e.g. "Name of the work: Laying of CC road … . ecv: 1723159.00". */
  workLine: string
  /** Participating bidders, in order — one column each (must be ≥ 1). */
  bidders: string[]
  /** ECV in rupees for the ECV row; null leaves it blank. */
  ecv: number | null
  /** Signature block for the bottom-right, e.g. "Executive Engineer,\nGajularamaram circle,CMC". */
  signature: string
}

// Row layout (1-based). Rows 5–15 hold the eleven criteria; the bid-capacity
// block follows.
const R_TITLE = 1
const R_NIT = 2
const R_WORK = 3
const R_HEADER = 4
const R_CRITERIA_START = 5 // Sl.No 1
const R_A = 16 // Sl.No 12 — Turnover (A)
const R_N = 17 // Sl.No 13 — N (period in years)
const R_B = 18 // Sl.No 14 — B (existing commitments)
const R_3ANB = 19 // Sl.No 15 — 3AN − B (bid capacity)
const R_ECV = 20 // Sl.No 16 — ECV
const R_REMARKS = 21 // Sl.No 17 — Remarks
const R_SURPLUS = 22 // Sl.No 18 — bid capacity − ECV
const R_SIGN = 26 // signature sits two rows below where it used to (was 24)

const FIRST_BIDDER_COL = 3 // column C
const ACCOUNTING_FMT = '_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_ '

/** 1-based column index → Excel column letter ("A", "B", … "AA"). */
export function columnLetter(n: number): string {
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

const THIN = { style: 'thin' as const }
const ALL_BORDERS = { top: THIN, left: THIN, bottom: THIN, right: THIN }

/**
 * Build the Evaluation Sheet workbook and return its .xlsx bytes.
 * Deterministic: the same input always produces the same sheet.
 */
export async function buildEvaluationSheet(input: EvaluationSheetInput): Promise<Buffer> {
  const bidders = input.bidders.map((b) => b.trim()).filter(Boolean)
  const nCols = 2 + Math.max(bidders.length, 1) // Sl.No + Requirements + one per bidder
  const lastColLetter = columnLetter(nCols)

  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Evaluation')

  // Column widths: Sl.No narrow, Requirements wide, bidder columns medium.
  ws.getColumn(1).width = 5
  ws.getColumn(2).width = 58
  for (let c = FIRST_BIDDER_COL; c <= nCols; c++) ws.getColumn(c).width = 20

  // Banner rows — each merged across the whole table width.
  const banner = (row: number, text: string, size: number) => {
    ws.mergeCells(`A${row}:${lastColLetter}${row}`)
    const cell = ws.getCell(row, 1)
    cell.value = text
    cell.font = { name: 'Book Antiqua', bold: true, size }
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    cell.border = ALL_BORDERS
  }
  banner(R_TITLE, 'TENDER DETAILS', 12)
  banner(R_NIT, input.nitLine, 11)
  banner(R_WORK, input.workLine, 11)

  // Header row: Sl. No | Requirements | <bidder names>.
  const headerCells = ['Sl. No.', 'Requirements as per Tender conditions', ...bidders]
  headerCells.forEach((text, i) => {
    const cell = ws.getCell(R_HEADER, i + 1)
    cell.value = text
    cell.font = { name: 'Book Antiqua', bold: true, size: 10 }
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
    cell.border = ALL_BORDERS
  })

  // A labelled data row: Sl.No in A, label in B, then one cell per bidder (its
  // value comes from `bidderValue`, a formula/number/blank). `boldLabel` bolds
  // the B label (the bid-capacity rows are bold in the office's sheet).
  const dataRow = (
    row: number,
    slNo: number | '',
    label: string,
    bidderValue: (col: number, letter: string) => ExcelJS.CellValue,
    opts: { boldLabel?: boolean; accounting?: boolean } = {}
  ) => {
    const a = ws.getCell(row, 1)
    a.value = slNo
    a.font = { name: 'Times New Roman', size: 10 }
    a.alignment = { horizontal: 'center', vertical: 'middle' }
    a.border = ALL_BORDERS

    const b = ws.getCell(row, 2)
    b.value = label
    b.font = { name: 'Times New Roman', bold: !!opts.boldLabel, size: 10 }
    b.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true }
    b.border = ALL_BORDERS

    for (let c = FIRST_BIDDER_COL; c <= nCols; c++) {
      const cell = ws.getCell(row, c)
      cell.value = bidderValue(c, columnLetter(c))
      cell.font = { name: 'Times New Roman', size: 10 }
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
      cell.border = ALL_BORDERS
      if (opts.accounting) cell.numFmt = ACCOUNTING_FMT
    }
  }

  // Eleven eligibility criteria — bidder cells blank for the office to fill.
  EVALUATION_CRITERIA.forEach((text, i) => dataRow(R_CRITERIA_START + i, i + 1, text, () => null))

  // Bid-capacity block. A / N / B are blank inputs; 3AN−B, Remarks and the
  // surplus are formulas that reference this bidder's own column.
  dataRow(R_A, 12, 'Turnover in Civil Engineering Works, Turnover Certificate issued by CA.(A)', () => null, {
    boldLabel: true,
    accounting: true
  })
  dataRow(R_N, 13, 'N', () => null, { boldLabel: true })
  dataRow(R_B, 14, 'B', () => null, { boldLabel: true, accounting: true })
  dataRow(R_3ANB, 15, '3AN-B', (_c, L) => ({ formula: `(3*${L}${R_A}*${L}${R_N})-${L}${R_B}` }), {
    boldLabel: true,
    accounting: true
  })
  dataRow(R_ECV, 16, 'ECV', () => (input.ecv != null ? input.ecv : null), { boldLabel: true, accounting: true })
  // Remarks are left blank under each bidder for the office to write by hand
  // (Responsive / Non-Responsive with the reason), not auto-derived.
  dataRow(R_REMARKS, 17, 'Remarks', () => null, { boldLabel: true })
  dataRow(R_SURPLUS, 18, '', (_c, L) => ({ formula: `${L}${R_3ANB}-${L}${R_ECV}` }), { accounting: true })

  // Signature, bottom-right under the last bidder column.
  const sign = ws.getCell(R_SIGN, nCols)
  sign.value = input.signature
  sign.font = { name: 'Book Antiqua', bold: true, size: 11 }
  sign.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }

  const out = await wb.xlsx.writeBuffer()
  return Buffer.from(out)
}

import * as ExcelJS from 'exceljs'
import { isReservedWork } from './tenderAgents/nameOfWork'

// Builds the "Evaluation Sheet" — the Bid Capacity evaluation the office runs on
// every participating bidder. One column per bidder, eleven fixed eligibility
// criteria, then the bid-capacity block: A (turnover), N (period, in years), B
// (existing commitments), 3AN−B (bid capacity), ECV and Remarks. The header,
// bidder columns and ECV are filled from the uploaded "View Bidders" sheet; the
// per-bidder data (turnover / N / B / eligibility / Remarks) is left blank for
// the office to key in from each bidder's documents, and the bid-capacity cell
// (3AN−B) plus the surplus are live formulas that resolve as those blanks fill.

/** Works with an estimate above this (rupees, 5 lakhs) are evaluated on 3AN−B. */
export const BID_CAPACITY_THRESHOLD = 500000
/** Indices in EVALUATION_CRITERIA that only apply to works evaluated on 3AN−B. */
const BID_CAPACITY_CRITERIA = [3, 4, 5]

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

/** Conditions that exist only for the 3AN−B bid-capacity check (commitments, the 3AN−B rule, CA turnover). */
export const isBidCapacityCondition = (c: string) =>
  /existing\s+commitments|administrative\s+sanction\s+amount|maximum\s+value\s+of\s+civil\s+engineering/i.test(c)

const RESERVATION_CRITERIA = [EVALUATION_CRITERIA[7], EVALUATION_CRITERIA[8]]
const isReservationCondition = (c: string) => /caste\s*certificate|reservation\s+work/i.test(c)

/**
 * The caste-certificate and one-work-per-agency conditions apply only to a work
 * reserved for a category (SC, ST, Waddera, … — see isReservedWork). Dropped
 * from the list for an open work; added (after the last condition before the
 * self-declaration, else at the end) when a reserved work's list lacks them,
 * e.g. a portal list fetched for it that doesn't carry them.
 */
export function applyReservationConditions(conditions: string[], reserved: boolean): string[] {
  if (!reserved) return conditions.filter((c) => !isReservationCondition(c))
  const out = [...conditions]
  const hasCaste = out.some((c) => /caste\s*certificate/i.test(c))
  const hasOneWork = out.some((c) => /reservation\s+work/i.test(c))
  const missing = [hasCaste ? null : RESERVATION_CRITERIA[0], hasOneWork ? null : RESERVATION_CRITERIA[1]].filter((c): c is string => c != null)
  if (missing.length === 0) return out
  const at = out.findIndex((c) => /self-?\s*declaration/i.test(c))
  out.splice(at >= 0 ? at : out.length, 0, ...missing)
  return out
}

/** Whether the 3AN−B bid-capacity block applies: estimate above 5 lakhs, or unknown (the ECV is never used in its place). */
export function needsBidCapacityFor(estimate: number | null | undefined): boolean {
  return estimate == null || estimate > BID_CAPACITY_THRESHOLD
}

/**
 * The conditions an evaluation sheet starts from: the tender's own list (portal
 * or saved) or the standard eleven, minus the bid-capacity conditions for a small
 * work, with the reservation conditions following the work name. The Evaluation
 * Sheet tab shows this list for the office to add to, remove from or reword.
 */
export function defaultConditions(tenderList: string[] | null | undefined, workLine: string, estimate: number | null | undefined): string[] {
  const needsBidCapacity = needsBidCapacityFor(estimate)
  const list = (tenderList ?? []).map((c) => c.trim()).filter(Boolean)
  const base = list.length > 0 ? list : needsBidCapacity ? [...EVALUATION_CRITERIA] : EVALUATION_CRITERIA.filter((_, i) => !BID_CAPACITY_CRITERIA.includes(i))
  return applyReservationConditions(
    // A portal list is filtered by wording too: it isn't in the standard order.
    base.filter((c) => needsBidCapacity || !isBidCapacityCondition(c)),
    isReservedWork(workLine)
  )
}

export interface EvaluationSheetInput {
  /** Full NIT line for row 2, e.g. "NIT No: 16/DB/EE/…/CMC/2026-27,Dt:25.07.2026,Tender ID:722231". */
  nitLine: string
  /** Full work line for row 3, e.g. "Name of the work: Laying of CC road … . ecv: 1723159.00". */
  workLine: string
  /** Participating bidders, in order — one column each (must be ≥ 1). */
  bidders: string[]
  /** ECV in rupees for the ECV row; null leaves it blank. */
  ecv: number | null
  /**
   * Estimate amount (rupees) from the Works List — a different figure from the
   * ECV. Decides whether the 3AN−B rows are needed; when unknown they are kept.
   */
  estimate?: number | null
  /**
   * Tender conditions for the requirement rows, as listed on the portal's
   * "Required Tender Documents Details" for this tender. When absent or empty,
   * the eleven standard conditions in EVALUATION_CRITERIA are used.
   */
  conditions?: string[]
  /**
   * The conditions were already settled (see defaultConditions) and possibly edited
   * by the office: print them exactly as given, applying no reservation / small-work
   * filtering of their own.
   */
  conditionsFinal?: boolean
  /** Signature block for the bottom-right, e.g. "Executive Engineer,\nGajularamaram circle,CMC". */
  signature: string
}

// Row layout (1-based). The criteria rows start at R_CRITERIA_START and run for
// as many conditions as the tender lists; the bid-capacity block follows them.
const R_TITLE = 1
const R_NIT = 2
const R_WORK = 3
const R_HEADER = 4
const R_CRITERIA_START = 5 // Sl.No 1

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

// Rough Excel-column-width-unit-to-character conversion for Book Antiqua
// 11pt, used only to pick a row height tall enough to show the wrapped Name
// of Work banner in full — same technique as estimateTemplate.ts's
// wrappedRowHeight, since exceljs never auto-fits a merged, wrapped cell's
// row height (and neither does every consuming spreadsheet application, e.g.
// Google Sheets), so it has to be computed and set explicitly here.
// Measured against a rendered sheet: 1.6 clipped the bold 11pt banner rows, so
// the estimates err on the tall side — a little spare space beats cut-off text.
const BANNER_CHARS_PER_WIDTH_UNIT = 1.0 // bold Book Antiqua 11pt
const CELL_CHARS_PER_WIDTH_UNIT = 1.05 // Times New Roman 10pt
const BANNER_LINE_POINTS = 15
const CELL_LINE_POINTS = 13

function wrappedRowHeight(text: string, widthUnits: number, charsPerUnit = BANNER_CHARS_PER_WIDTH_UNIT, linePoints = BANNER_LINE_POINTS): number {
  const charsPerLine = Math.max(10, Math.floor(widthUnits * charsPerUnit))
  // Wrap on words, as the spreadsheet does, instead of dividing the length — a
  // long word pushed to the next line costs a line the plain division misses.
  let lines = 1
  let used = 0
  for (const word of text.split(/\s+/).filter(Boolean)) {
    const w = word.length
    if (used > 0 && used + 1 + w > charsPerLine) {
      lines += 1
      used = w
    } else {
      used += (used > 0 ? 1 : 0) + w
    }
    while (used > charsPerLine) {
      lines += 1
      used -= charsPerLine
    }
  }
  return lines * linePoints + 4
}

/**
 * Build the Evaluation Sheet workbook and return its .xlsx bytes.
 * Deterministic: the same input always produces the same sheet.
 */
export async function buildEvaluationSheet(input: EvaluationSheetInput): Promise<Buffer> {
  const bidders = input.bidders.map((b) => b.trim()).filter(Boolean)
  const conditions = (input.conditions ?? []).map((c) => c.trim()).filter(Boolean)
  // Works with an estimate above 5 lakhs are evaluated on bid capacity (3AN−B); at
  // or below it the A / N / B / 3AN−B / surplus rows are dropped. With no estimate
  // the rows are kept (the ECV is never used in its place).
  const needsBidCapacity = needsBidCapacityFor(input.estimate)
  const reserved = isReservedWork(input.workLine)
  // Reservation is read from the work's name ("… (Reserved for SC)"), so the caste
  // certificate / one-work-per-agency conditions follow it unless already settled.
  const criteria = input.conditionsFinal ? conditions : defaultConditions(conditions, input.workLine, input.estimate)
  const nCrit = criteria.length
  // Rows after the criteria, in order; the bid-capacity ones only when needed.
  let next = R_CRITERIA_START + nCrit
  let sl = nCrit + 1
  const R_A = needsBidCapacity ? next++ : -1 // Turnover (A)
  const R_N = needsBidCapacity ? next++ : -1 // N (period in years)
  const R_B = needsBidCapacity ? next++ : -1 // B (existing commitments)
  const R_3ANB = needsBidCapacity ? next++ : -1 // 3AN − B (bid capacity)
  const R_ECV = next++ // ECV
  const R_REMARKS = next++ // Remarks
  const R_SURPLUS = needsBidCapacity ? next++ : -1 // bid capacity − ECV
  const R_SIGN = next + 3 // signature, below the last row
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
  // Name of Work is often long enough to wrap several lines — size the row
  // to fit it instead of leaving exceljs's default height, which clips it.
  const tableWidthUnits = 5 + 58 + (nCols - 2) * 20
  ws.getRow(R_WORK).height = wrappedRowHeight(input.workLine, tableWidthUnits)
  ws.getRow(R_NIT).height = wrappedRowHeight(input.nitLine, tableWidthUnits)

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

  // Eligibility criteria — bidder cells blank for the office to fill.
  criteria.forEach((text, i) => {
    dataRow(R_CRITERIA_START + i, i + 1, text, () => null)
    // A long condition wraps inside column B; size its row to show every line.
    ws.getRow(R_CRITERIA_START + i).height = Math.max(
      18,
      wrappedRowHeight(text, 58, CELL_CHARS_PER_WIDTH_UNIT, CELL_LINE_POINTS)
    )
  })

  // Bid-capacity block. A / N / B are blank inputs; 3AN−B, Remarks and the
  // surplus are formulas that reference this bidder's own column.
  if (needsBidCapacity) {
    dataRow(R_A, sl++, 'Turnover in Civil Engineering Works, Turnover Certificate issued by CA.(A)', () => null, {
      boldLabel: true,
      accounting: true
    })
    dataRow(R_N, sl++, 'N', () => null, { boldLabel: true })
    dataRow(R_B, sl++, 'B', () => null, { boldLabel: true, accounting: true })
    dataRow(R_3ANB, sl++, '3AN-B', (_c, L) => ({ formula: `(3*${L}${R_A}*${L}${R_N})-${L}${R_B}` }), {
      boldLabel: true,
      accounting: true
    })
  }
  dataRow(R_ECV, sl++, 'ECV', () => (input.ecv != null ? input.ecv : null), { boldLabel: true, accounting: true })
  // Remarks are left blank under each bidder for the office to write by hand
  // (Responsive / Non-Responsive with the reason), not auto-derived.
  dataRow(R_REMARKS, sl++, 'Remarks', () => null, { boldLabel: true })
  if (needsBidCapacity) {
    dataRow(R_SURPLUS, sl, '', (_c, L) => ({ formula: `${L}${R_3ANB}-${L}${R_ECV}` }), { accounting: true })
  }

  // Signature, bottom-right under the last bidder column.
  const sign = ws.getCell(R_SIGN, nCols)
  sign.value = input.signature
  sign.font = { name: 'Book Antiqua', bold: true, size: 11 }
  sign.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }

  const out = await wb.xlsx.writeBuffer()
  return Buffer.from(out)
}

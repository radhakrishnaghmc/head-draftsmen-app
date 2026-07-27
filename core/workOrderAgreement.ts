// Fills the bundled Work Order and Agreement formats (both .docx mail-merge
// templates) for one work. Both documents describe the same award, so they
// share a single set of canonical, editable fields (raw amounts + dates +
// party details); each template then formats those fields its own way — the
// Work Order prints plain "Rs. 415646.00" figures, while the Agreement prints
// Indian-grouped "Rs.19,93,085.00" figures and spells the contract value and
// its date out in words. Sources, in priority order, mirror Give Intimation:
// the portal "View Intimation Notice" HTML -> the L-1 selection / evaluation
// PDF -> the picked Works List row.
import { computeWorkAmounts, indianDigitGroups } from './worksAmounts'
import { amountToWords, dateToWords } from './numberToWords'
import type { IntimationNotice } from './intimationNotice'
import type { TenderEvaluation } from './tenderEvaluationPdf'

/** The canonical, editable inputs both documents are built from. */
export interface WorkOrderAgreementFields {
  circle: string
  cno: string
  zone: string
  nameOfWork: string
  agencyName: string
  address: string
  phone: string
  wincode: string
  financialYear: string
  /** "Amount of estimate" in Lakhs, as entered on the Works List (e.g. "25" or "0.95"). */
  estimateLakhs: string
  /** ECV in whole rupees, or "" when not yet known. */
  ecvRupees: string
  /** Tender percentage as a plain number string ("16.5"), or "" for nomination/unknown. */
  tenderPercent: string
  /** Contract amount in rupees, or "" when not yet known. */
  contractRupees: string
  /** Work Order issue date, "dd.mm.yyyy" — left blank if unknown (the office often hand-dates it). */
  workOrderDate: string
  /** Agreement execution date, "dd.mm.yyyy". */
  agreementDate: string
  /** "Note approved by the Zonal Commissioner Dt." reference date on the Work Order, "dd.mm.yyyy". */
  adminSanctionDate: string
}

/** Indian financial year for a date (1 April boundary): 2026-07-27 -> "2026-27". */
export function indianFinancialYear(d = new Date()): string {
  const y = d.getFullYear()
  const startY = d.getMonth() >= 3 ? y : y - 1
  return `${startY}-${String((startY + 1) % 100).padStart(2, '0')}`
}

function num(v: string | number | null | undefined): number | null {
  if (v == null || v === '') return null
  const n = Number(String(v).replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : null
}

/** Plain 2-decimal figure, no digit grouping — the Work Order's own "415646.00" style. */
function money2(n: number): string {
  return n.toFixed(2)
}

/** "Rs.19,93,085.00" — the Agreement's Indian-grouped style. */
function groupedRupees(n: number): string {
  const [whole, frac] = n.toFixed(2).split('.')
  return `Rs.${indianDigitGroups(Number(whole))}.${frac}`
}

/** Tender percentage to the office's fixed 2-decimal wording ("16.50", "27.45"). */
function formatPercent(pct: number): string {
  return pct.toFixed(2)
}

/**
 * Pre-fill the canonical fields from whatever's been uploaded/picked. Every
 * field is still user-editable afterwards; the two date fields the office
 * hand-writes (Agreement date, Admin sanction date) default to the L-1
 * selection date when the PDF carries one, else blank.
 */
export function deriveFields(
  notice: IntimationNotice,
  pdf: TenderEvaluation,
  row: Record<string, string>
): WorkOrderAgreementFields {
  const est = computeWorkAmounts(row)
  const ecv = notice.ecvRupees ?? pdf.ecvRupees ?? est.ecv ?? null
  const tenderPct = pdf.tenderPercentage ?? num(row['Tender Percentage'])
  // Keep paise: the office prints the contract value to 2 decimals
  // ("Rs.14,45,983.17"), so the computed fallback is rounded to paise, not
  // to whole rupees the way the Works List's Contract Amount column is.
  const contract =
    notice.contractRupees ??
    pdf.contractRupees ??
    (ecv != null && tenderPct != null ? Math.round(ecv * (1 - tenderPct / 100) * 100) / 100 : null)
  const loaDate = pdf.noticeDate ?? ''

  return {
    circle: row['Circle'] ?? '',
    cno: row['CNO'] ?? '',
    zone: row['Zone'] ?? '',
    nameOfWork: row['Name of the work'] ?? '',
    agencyName: notice.agencyName ?? pdf.l1AgencyName ?? row['Name of the Agency'] ?? '',
    address: notice.address ?? row['Address of the agency'] ?? '',
    phone: row['Phone number of the agency'] ?? '',
    wincode: row['Wincode'] ?? '',
    financialYear: indianFinancialYear(),
    estimateLakhs: (row['Amount of estimate'] ?? '').replace(/,/g, '').trim(),
    ecvRupees: ecv != null ? String(ecv) : '',
    tenderPercent: tenderPct != null ? String(tenderPct) : '',
    contractRupees: contract != null ? String(contract) : '',
    workOrderDate: loaDate,
    agreementDate: loaDate,
    adminSanctionDate: ''
  }
}

/** The {{Label}} -> value map for the Work Order template. */
export function workOrderPlaceholders(f: WorkOrderAgreementFields): Record<string, string> {
  const estLakhs = num(f.estimateLakhs)
  const ecv = num(f.ecvRupees)
  const pct = num(f.tenderPercent)
  const contract = num(f.contractRupees)
  return {
    Circle: f.circle,
    CNO: f.cno,
    Financialyear: f.financialYear,
    'Agreement date': f.workOrderDate,
    'Name of the agency': f.agencyName,
    'Address of the agency': f.address,
    'Phone no.': f.phone,
    Zone: f.zone,
    'Name of the work': f.nameOfWork,
    'Administrative Sanction date': f.adminSanctionDate,
    Wincode: f.wincode,
    'Estimate Amount': estLakhs != null ? `Rs. ${estLakhs.toFixed(2)} Lakhs` : '',
    ECV: ecv != null ? `Rs. ${money2(ecv)}` : '',
    'Tender Percentage': pct == null ? '' : pct === 0 ? '0%' : `(-) ${formatPercent(pct)}%-Less`,
    'Contract Amount': contract != null ? `Rs. ${money2(contract)}` : '',
    TP: pct != null ? formatPercent(pct) : ''
  }
}

/** The {{Label}} -> value map for the Agreement template. */
export function agreementPlaceholders(f: WorkOrderAgreementFields): Record<string, string> {
  const estLakhs = num(f.estimateLakhs)
  const ecv = num(f.ecvRupees)
  const pct = num(f.tenderPercent)
  const contract = num(f.contractRupees)
  return {
    Circle: f.circle,
    CNO: f.cno,
    'Name of the work': f.nameOfWork,
    // Template already prints "Rs." … "Lakhs" around this one.
    'Estimate Amount': estLakhs != null ? estLakhs.toFixed(2) : '',
    ECV: ecv != null ? groupedRupees(ecv) : '',
    // Template already prints "(-)" … "%-Less" around this one.
    'Tender percentage': pct != null ? formatPercent(pct) : '',
    'Contract value': contract != null ? groupedRupees(contract) : '',
    'Agreement Date': f.agreementDate,
    'Agreement date in words': dateToWords(f.agreementDate),
    'Agency Name': f.agencyName,
    'Contract value in rupees': contract != null ? amountToWords(contract) : ''
  }
}

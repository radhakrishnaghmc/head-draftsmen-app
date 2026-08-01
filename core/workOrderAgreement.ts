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

// The agency's postal address prints in a narrow "To," block, under the agency
// name — so it must wrap to short lines instead of running the full page width.
const ADDRESS_WRAP_WIDTH = 21

/**
 * Wrap an agency address to lines no wider than `maxWidth` characters (spaces
 * included), breaking at spaces and hard-splitting any single token longer than
 * the width, joined with "\n". The docx fill turns each "\n" into a real Word
 * line break (see core/docx-edit.ts's setRunText), so the address sits in a tidy
 * narrow block under the agency name rather than spanning the whole line.
 */
export function wrapAgencyAddress(address: string, maxWidth: number = ADDRESS_WRAP_WIDTH): string {
  const clean = (address ?? '').replace(/\s+/g, ' ').trim()
  if (!clean) return ''
  const lines: string[] = []
  let cur = ''
  for (const rawWord of clean.split(' ')) {
    let word = rawWord
    // Hard-break a token that can't fit on a line on its own.
    while (word.length > maxWidth) {
      if (cur) {
        lines.push(cur)
        cur = ''
      }
      lines.push(word.slice(0, maxWidth))
      word = word.slice(maxWidth)
    }
    if (!cur) cur = word
    else if (cur.length + 1 + word.length <= maxWidth) cur += ` ${word}`
    else {
      lines.push(cur)
      cur = word
    }
  }
  if (cur) lines.push(cur)
  return lines.join('\n')
}

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
  /** Corporation abbreviation for the Forwarding Slip (e.g. "CMC"), from the chosen office. */
  corporation: string
  /** Corporation full name for the Forwarding Slip title (e.g. "Cyberabad Municipal Corporation"). */
  corporationFullName: string
  /** Technical Sanction No & Date, hand-entered for the Forwarding Slip (e.g. "11/26-27, Dt: 29.05.2026"). */
  tsNoDate: string
  /** Period of completion in months, hand-entered for the Forwarding Slip (e.g. "02"). */
  completionMonths: string
  /** Reservation category on the work ("SC" / "ST" / ""), for the Forwarding Slip's EMD/ASD exemption. */
  reservation: string
}

/** SC/ST reservation drives EMD/ASD exemption — read from the work name's "(Reserved to SC/ST)" tag or a Reservation column. */
export function reservationFromRow(row: Record<string, string>): string {
  const explicit = (row['Reservation'] ?? '').trim()
  if (/\b(SC|ST)\b/i.test(explicit)) return explicit.toUpperCase().match(/\b(SC|ST)\b/i)![0].toUpperCase()
  const m = /reserved\s*(?:to|for)?\s*(SC|ST)\b/i.exec(row['Name of the work'] ?? '')
  return m ? m[1].toUpperCase() : ''
}

/** The Circle name and number out of a NIT No ("…/EE/Gajularamaram Circle-57/QBZ/CMC/…" -> {circle:"Gajularamaram", cno:"57"}). */
export function circleFromNit(nit: string | undefined): { circle: string; cno: string } {
  const m = /\/EE\/\s*(.+?)\s*Circle-\s*(\d+)/i.exec(nit ?? '')
  return m ? { circle: m[1].trim(), cno: m[2] } : { circle: '', cno: '' }
}

/**
 * A synthetic Works-List-shaped row built purely from an uploaded L-1
 * evaluation + Online Intimation, for the Tools workspace's standalone Work
 * Order / Agreement generators — which are deliberately NOT tied to the Works
 * List and do no Zone/Circle verification. Circle & number come from the NIT
 * No, the work name from the L-1 sheet; every other column is left for
 * deriveFields to fill from the uploads (agency/address/amounts/dates).
 */
export function standaloneRowFromSources(pdf: TenderEvaluation, notice: IntimationNotice): Record<string, string> {
  const { circle, cno } = circleFromNit(pdf.noticeNo || notice.nitNo)
  return {
    Circle: circle,
    'Circle number': cno,
    Zone: '',
    'Name of the work': pdf.nameOfWork ?? '',
    // The tender Notice No (= NIT No) and its date, from the uploaded L-1 sheet
    // (falling back to the Intimation) — so the Note Submitted's "NIT No:" fills
    // here just like it would from a Works List row.
    'Tender Notice No': pdf.noticeNo || notice.nitNo || '',
    'Tender notice Date': pdf.noticeDate || ''
  }
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
  // The name of work comes from the uploaded L-1 sheet — the Works List row only
  // supplies supporting details (Circle/CNO/estimate/…) when its name matched.
  const workName = pdf.nameOfWork || row['Name of the work'] || ''

  return {
    circle: row['Circle'] ?? '',
    cno: row['Circle number'] ?? row['CNO'] ?? '',
    zone: row['Zone'] ?? '',
    nameOfWork: workName,
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
    adminSanctionDate: '',
    corporation: '',
    corporationFullName: '',
    tsNoDate: '',
    completionMonths: '',
    reservation: reservationFromRow({ ...row, 'Name of the work': workName })
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
    'Address of the agency': wrapAgencyAddress(f.address),
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

/** The {{Label}} -> value map for the Forwarding Slip template. */
export function forwardingSlipPlaceholders(f: WorkOrderAgreementFields): Record<string, string> {
  const estLakhs = num(f.estimateLakhs)
  const ecv = num(f.ecvRupees)
  const pct = num(f.tenderPercent)
  const contract = num(f.contractRupees)
  const reserved = /\b(SC|ST)\b/i.test(f.reservation)
  const EXEMPT = 'Exempted for reservation work'
  // Contractor block: name, address, phone — each on its own line.
  const contractor = [f.agencyName, f.address, f.phone].map((s) => (s ?? '').trim()).filter(Boolean).join('\n')
  return {
    Corporation: f.corporation,
    'Corporation Full Name': f.corporationFullName.toUpperCase(),
    Circle: f.circle,
    CNO: f.cno,
    Financialyear: f.financialYear,
    'Name of the work': f.nameOfWork,
    'Amount of Estimation': estLakhs != null ? groupedRupees(estLakhs * 100000) : '',
    'Contractor Name and Address': contractor,
    'Approximate Value': contract != null ? groupedRupees(contract) : '',
    'Tender Percentage': pct == null ? '' : pct === 0 ? '0%' : `(-) ${formatPercent(pct)}% Less`,
    'Technical Sanction No and Date': f.tsNoDate,
    'Period of completion': f.completionMonths.trim() ? `${f.completionMonths.trim()}   Months` : '',
    // Reserved works are EMD-exempt; otherwise the 1% / 1.5% amounts off ECV.
    'EMD 1 percent': reserved ? EXEMPT : ecv != null ? groupedRupees(ecv * 0.01) : '',
    'EMD 1.5 percent': reserved ? EXEMPT : ecv != null ? groupedRupees(ecv * 0.015) : '',
    // ASD applies only when the tender % exceeds 25 (at (%-25)% of ECV); else "-".
    'ASD Details': !reserved && pct != null && pct > 25 && ecv != null ? groupedRupees(ecv * ((pct - 25) / 100)) : '-'
  }
}

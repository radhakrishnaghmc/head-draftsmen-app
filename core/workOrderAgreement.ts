// Fills the bundled Work Order and Agreement formats (both .docx mail-merge
// templates) for one work. Both documents describe the same award, so they
// share a single set of canonical, editable fields (raw amounts + dates +
// party details); each template then formats those fields its own way — the
// Work Order prints plain "Rs. 415646.00" figures, while the Agreement prints
// Indian-grouped "Rs.19,93,085.00" figures and spells the contract value and
// its date out in words. Sources, in priority order, mirror Give Intimation:
// the portal "View Intimation Notice" HTML -> the L-1 selection / evaluation
// PDF -> the picked Works List row.
import { computeWorkAmounts, formatRupees, indianDigitGroups, tenderPercentFromRow } from './worksAmounts'
import { amountToWords, dateToWords, integerToIndianWords } from './numberToWords'
import { zoneAbbr } from './loaSe'
import { extractItemNo, stripItemNoTag } from './bidDocument'
import type { IntimationNotice } from './intimationNotice'
import type { TenderEvaluation } from './tenderEvaluationPdf'
import type { BalanceEmdReceipt } from './balanceEmdReceipt'

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
  /** The NIT's item number (e.g. "3"), extracted from a "(Item No.3)" tag in the work name — SE offices tender multiple items under one NIT number. "" when no such tag is present. */
  itemNo: string
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

const DIGIT_WORDS = ['Zero', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine']

/**
 * A tender percentage spelled out for the Price Bid's "(In words)" blank —
 * integer part in words, "Point", then each decimal digit named:
 * 16.5 -> "Sixteen Point Five Zero Percent", 27.45 -> "Twenty Seven Point Four
 * Five Percent". (The 2-decimal form matches the office's "%" figures.)
 */
export function percentToWords(pct: number): string {
  if (!Number.isFinite(pct)) return ''
  const [intPart, decPart] = Math.abs(pct).toFixed(2).split('.')
  const intWords = integerToIndianWords(Number(intPart))
  const decWords = decPart.split('').map((d) => DIGIT_WORDS[Number(d)]).join(' ')
  return `${intWords} Point ${decWords} Percent`
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
  const tenderPct = pdf.tenderPercentage ?? num(tenderPercentFromRow(row))
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
  // SE offices tag the item number onto the work name itself ("...(Item No.3)")
  // — pull it into its own field and strip the tag back out, so it doesn't also
  // print as part of the printed work name.
  const itemNo = extractItemNo(workName) ?? ''

  return {
    circle: row['Circle'] ?? '',
    cno: row['Circle number'] ?? row['CNO'] ?? '',
    zone: row['Zone'] ?? '',
    nameOfWork: stripItemNoTag(workName),
    itemNo,
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

// Hand-writable blank date lines — printed when no agreement/work-order date has
// been set, so the document shows the "Dt:" / "Date:" label followed by a ruled
// blank to fill in by hand (matching the templates' own "____" convention),
// instead of silently defaulting to the tender-notice date.
const DATE_BLANK = '____________'
const DATE_WORDS_BLANK = '______________________'

/** The {{Label}} -> value map for the Work Order template. */
export function workOrderPlaceholders(f: WorkOrderAgreementFields): Record<string, string> {
  const estLakhs = num(f.estimateLakhs)
  const ecv = num(f.ecvRupees)
  const pct = num(f.tenderPercent)
  const contract = num(f.contractRupees)
  return {
    Circle: f.circle,
    CNO: f.cno,
    // Corporation abbreviation for the EE sign-off block (was a hard-coded "CMC"
    // in the template — now the work's actual corporation).
    Corporation: f.corporation,
    Financialyear: f.financialYear,
    'Agreement date': f.workOrderDate.trim() || DATE_BLANK,
    'Name of the agency': f.agencyName,
    'Address of the agency': wrapAgencyAddress(f.address),
    'Phone no.': f.phone,
    Zone: f.zone,
    // Zone abbreviation for the Work Order number / Zonal-Commissioner ref (was a
    // hard-coded "QBZ" in the template — now reflects the work's actual zone).
    ZoneAbbr: zoneAbbr(f.zone),
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
    'Agreement Date': f.agreementDate.trim() || DATE_BLANK,
    'Agreement date in words': dateToWords(f.agreementDate) || DATE_WORDS_BLANK,
    'Agency Name': f.agencyName,
    'Contract value in rupees': contract != null ? amountToWords(contract) : ''
  }
}

/**
 * The {{Label}} -> value map for the File Backer — the file's cover page. A
 * centred cover: the corporation header (fixed in the template) over the circle
 * line ("{{Circle Header}}" = "GAJULARAMARAM CIRCLE-57") and a six-row table of
 * the work's key facts. Amounts follow the same wording as the Work Order /
 * Agreement (estimate in Lakhs, contract Indian-grouped, "% Less").
 */
export function fileBackerPlaceholders(f: WorkOrderAgreementFields): Record<string, string> {
  const estLakhs = num(f.estimateLakhs)
  const ecv = num(f.ecvRupees)
  const pct = num(f.tenderPercent)
  const contract = num(f.contractRupees)
  const circleHeader = f.circle ? `${f.circle.toUpperCase()} CIRCLE${f.cno ? `-${f.cno}` : ''}` : ''
  return {
    'Circle Header': circleHeader,
    'Name of the work': f.nameOfWork,
    'Estimate Amount': estLakhs != null ? `Rs. ${estLakhs.toFixed(2)} Lakhs` : '',
    // ECV sits between Estimate and Contract; it's stored in rupees (blank stays
    // blank — never derived from the estimate), grouped like the Contract Amount.
    ECV: ecv != null ? groupedRupees(ecv) : '',
    'Contract Amount': contract != null ? groupedRupees(contract) : '',
    'Tender Percentage': pct == null ? '' : pct === 0 ? 'At par' : `${formatPercent(pct)} % Less`,
    'Name of the agency': f.agencyName,
    'Address of the agency': f.address
  }
}

/**
 * The {{Label}} -> value map for the QCC (Quality Control Cell) Intimation
 * letter — the Dy.EE's request to the Quality Control Division to inspect a
 * work about to start. Reuses the same office/work/agency/amount data as the
 * Agreement; every other field on the form (Lr.No, dates, GST No., Date of Mark
 * Out, Status of Work, and the four Yes/No document checks) is hand-filled and
 * carries no placeholder. The same template is ALSO offered on the Issue
 * Documents tab, where it's filled by column-header matching + office bake — so
 * these placeholder names deliberately mirror the Works List column headers and
 * the office tokens ({{Corporation Full Name}}/{{Corporation}}/{{Circle}}/{{CNO}}),
 * and amounts use the same formatRupees "Rs …/-" form withComputedAmounts emits.
 */
export function qccIntimationPlaceholders(f: WorkOrderAgreementFields): Record<string, string> {
  const estLakhs = num(f.estimateLakhs)
  const ecv = num(f.ecvRupees)
  const contract = num(f.contractRupees)
  return {
    'Corporation Full Name': f.corporationFullName.toUpperCase(),
    Corporation: f.corporation,
    Circle: f.circle,
    CNO: f.cno,
    'Name of the work': f.nameOfWork,
    'Amount of estimate': estLakhs != null ? formatRupees(estLakhs * 100000) : '',
    ECV: ecv != null ? formatRupees(ecv) : '',
    'Contract Amount': contract != null ? formatRupees(contract) : '',
    Wincode: f.wincode,
    'Name of the Agency': f.agencyName,
    'Completion Period': f.completionMonths.trim() ? `${f.completionMonths.trim()} Months` : ''
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
    // Answers the slip's "Are the rates … within the Estimate rates …?" — so the
    // quoted % is followed by ", within the estimate rates" (at par or below the
    // estimate is within-estimate; blank stays blank).
    'Tender Percentage':
      pct == null ? '' : pct === 0 ? '0%, within the estimate rates' : `(-) ${formatPercent(pct)}% Less, within the estimate rates`,
    'Technical Sanction No and Date': f.tsNoDate,
    'Period of completion': f.completionMonths.trim() ? `${f.completionMonths.trim()}   Months` : '',
    // Reserved works are EMD-exempt; otherwise the 1% / 1.5% amounts off ECV.
    'EMD 1 percent': reserved ? EXEMPT : ecv != null ? groupedRupees(ecv * 0.01) : '',
    'EMD 1.5 percent': reserved ? EXEMPT : ecv != null ? groupedRupees(ecv * 0.015) : '',
    // ASD applies only when the tender % exceeds 25 (at (%-25)% of ECV); else "-".
    'ASD Details': !reserved && pct != null && pct > 25 && ecv != null ? groupedRupees(ecv * ((pct - 25) / 100)) : '-'
  }
}

/**
 * Placeholders for the full Civil Tender Document (the 41-page NIT/tender
 * document, page 1 = the Forwarding Slip). Work fields come from `f` (Works List
 * row + L-1); the NIT No, the four bid dates (Downloading Start/End, Receipt of
 * Bids, Price Bid Opening) and the L-1 agency come from the uploaded L-1 sheet
 * (`pdf`); "No. of pages of Agreement" and "No. of items in Schedule 'A'" the
 * office hand-enters (`extras`).
 */
export function civilTenderPlaceholders(
  f: WorkOrderAgreementFields,
  pdf: TenderEvaluation,
  extras: { pagesOfAgreement?: string; scheduleAItems?: string }
): Record<string, string> {
  const estLakhs = num(f.estimateLakhs)
  const ecv = num(f.ecvRupees)
  const pct = num(f.tenderPercent)
  const contract = num(f.contractRupees)
  const contractor = [f.agencyName, f.address].map((s) => (s ?? '').trim()).filter(Boolean).join('\n')
  return {
    'Name of the work': f.nameOfWork,
    'Estimate Amount': estLakhs != null ? groupedRupees(estLakhs * 100000) : '',
    // Page-1 forwarding slip answers "Are the rates … within the Estimate
    // rates …?" — same wording as the standalone Forwarding Slip.
    'Tender Percentage':
      pct == null ? '' : pct === 0 ? '0%, within the estimate rates' : `(-) ${formatPercent(pct)}% Less, within the estimate rates`,
    Contractor: contractor,
    'Contract Amount': contract != null ? groupedRupees(contract) : '',
    ECV: ecv != null ? groupedRupees(ecv) : '',
    EMD: ecv != null ? groupedRupees(Math.round(ecv * 0.01)) : '',
    'NIT No': pdf.noticeNo ?? '',
    Period: f.completionMonths.trim(),
    // Bid document downloading start/end; receipt of bids = closing.
    'Bid Start': pdf.bidStart ?? '',
    'Bid End': pdf.bidClose ?? '',
    'Bid Receipt': pdf.bidClose ?? '',
    // Price bid opening = the sheet's Server Time (when the price bid was opened).
    'Price Bid Open': pdf.serverDate ?? '',
    'Pages Of Agreement': extras.pagesOfAgreement ?? '',
    'Schedule A Items': extras.scheduleAItems ?? '',
    // Office block for the page footer (repeats on every page).
    Circle: f.circle,
    CNO: f.cno,
    Zone: f.zone,
    // Price Bid page: the "I Sri/Smt/M/s …" agency, the estimated contract value
    // in figures + words, and the tender % in figures and words.
    'Agency Name': f.agencyName,
    'ECV Figures Words': ecv != null ? `${groupedRupees(ecv)} (${amountToWords(ecv)})` : '',
    'Tender Percentage Figures': pct == null ? '' : `${formatPercent(pct)} %`,
    'Tender Percentage Words': pct == null ? '' : percentToWords(pct),
    // "Details of Maximum amount reimbursable to the Contractor" — GST @ 18% and
    // Labour Cess @ 1%, both off the ECV.
    GST: ecv != null ? groupedRupees(ecv * 0.18) : '',
    'Labour Cess': ecv != null ? groupedRupees(ecv * 0.01) : ''
  }
}

/** Indian-grouped rupees with paise, WITHOUT the "Rs." prefix (the zonal templates print "Rs." themselves). */
function groupedAmount(n: number): string {
  const [whole, frac] = n.toFixed(2).split('.')
  return `${indianDigitGroups(Number(whole))}.${frac}`
}

/**
 * Placeholders for the three Zone-level (SE office) documents — the Work Order,
 * the Memo Concluding Agreement, and the Memo forwarding the Agreement Bond to
 * the EE. One combined map serves all three: each template's fill only consumes
 * the labels it actually contains.
 *
 * The office is a Zone (no Circle of its own), so the Executive Engineer /
 * "{{EE Circle}}" named throughout is the *work's* own Circle, taken from the
 * derived fields (NIT / directory), never the office. Serial numbers, dates,
 * page counts and the EMD 1% Payment ID/UTR/CMS references (a separate, bank-
 * transferred payment the uploaded Balance EMD receipt doesn't cover) are left
 * blank in the templates for the office to hand-write.
 */
export function zonalDocsPlaceholders(
  f: WorkOrderAgreementFields,
  notice: IntimationNotice,
  pdf: TenderEvaluation,
  emd: BalanceEmdReceipt = {}
): Record<string, string> {
  const ecv = num(f.ecvRupees)
  const contract = num(f.contractRupees)
  const estLakhs = num(f.estimateLakhs)
  const zone = (f.zone ?? '').trim()
  const eeCircle = f.circle ? `${f.circle} Circle${f.cno ? `-${f.cno}` : ''}` : ''
  return {
    Zone: zone,
    'Zone Caps': zone ? `${zone.toUpperCase()} ZONE` : '',
    ZoneAbbr: zoneAbbr(zone),
    Corp: f.corporation,
    'Corp Full': (f.corporationFullName ?? '').toUpperCase(),
    FY: f.financialYear,
    'EE Circle': eeCircle,
    'Agency Name': f.agencyName,
    // Wrap the address across lines (as the circle-level Work Order does) so it
    // doesn't run off in one long line — the fill turns the \n into hard breaks.
    'Agency Address': wrapAgencyAddress(f.address),
    'Agency Phone': f.phone,
    'Name of Work': f.nameOfWork,
    'Item No': f.itemNo,
    'NIT No': notice.nitNo || pdf.noticeNo || '',
    'NIT Date': notice.nitDate || pdf.noticeDate || '',
    'Tender ID': pdf.tenderId ?? '',
    'Estimate Lakhs': estLakhs != null ? estLakhs.toFixed(2) : f.estimateLakhs,
    // The Agreement Bond paper prints the estimate as a full grouped rupee
    // figure ("Rs. 3,53,00,000.00"), not "X Lakhs" like the other SE docs.
    'Estimate Rupees': estLakhs != null ? groupedAmount(estLakhs * 100000) : '',
    ECV: ecv != null ? groupedAmount(ecv) : '',
    Contract: contract != null ? groupedAmount(contract) : '',
    // The Contract Deed's consideration clause spells the contract value out
    // in words, the same way the circle-level Agreement Bond does.
    'Contract Words': contract != null ? amountToWords(contract) : '',
    'Tender Percent': f.tenderPercent,
    Period: f.completionMonths.trim(),
    // The uploaded Balance EMD receipt gives the amount actually paid — more
    // reliable than a re-derived 1.5% guess (which, unlike the receipt, never
    // accounts for the 2.5% reserved-work rate) — so prefer it when present.
    'EMD 1.5%': emd.balanceEmdRupees != null ? indianDigitGroups(Math.round(emd.balanceEmdRupees)) : ecv != null ? indianDigitGroups(Math.round(ecv * 0.015)) : '',
    'EMD 1%': ecv != null ? indianDigitGroups(Math.round(ecv * 0.01)) : '',
    // e-Corpus Fund contribution to the Managing Director, TSTS — 0.04% of ECV
    // (matches the SE Letter of Acceptance's own e-Corpus derivation).
    'E-Corpus': ecv != null ? indianDigitGroups(Math.round(ecv * 0.0004)) : '',
    'EMD Receipt No': emd.receiptNo ?? '',
    'EMD Receipt Date': emd.paymentDate ?? '',
    // The Agreement Bond's own execution date — same shared date picked for the
    // EE Work Order / Agreement Bond (f.agreementDate), left blank the same way
    // when unset rather than guessing at a date.
    'Agreement Date': f.agreementDate.trim() || DATE_BLANK,
    'Agreement date in words': dateToWords(f.agreementDate) || DATE_WORDS_BLANK,
    // The zonal Work Order's own "DT." line — same shared date as the
    // Agreement Bond above, left blank the same way when unset.
    'Work Order Date': f.workOrderDate.trim() || DATE_BLANK
  }
}

import { computeWorkAmounts, tenderPercentFromRow } from './worksAmounts'
import { wrapAgencyAddress, type WorkOrderAgreementFields } from './workOrderAgreement'
import { isReservedWork } from './tenderAgents/nameOfWork'
import { zoneAbbr } from './loaSe'
import type { IntimationNotice } from './intimationNotice'
import type { TenderEvaluation } from './tenderEvaluationPdf'

/**
 * The office details an Intimation letter reads from the chosen office when the
 * work's own row doesn't carry them — a structural subset of src/office.ts's
 * `Office`, kept here so this core module doesn't reach into the renderer.
 */
export interface IntimationOffice {
  circle?: string
  circleNumber?: string
  zone?: string
  corporation?: string
  corporationFullName?: string
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

/**
 * EMD @ 1.5% is exempted for works reserved for a particular category — the
 * work name carries a "reserved for SC/ST/Waddera/Vaddera/WLCCS/…" tag. Any
 * category counts (core/tenderAgents/nameOfWork.ts's isReservedWork, the
 * single shared detector core/workOrderAgreement.ts and core/noteSubmitted.ts
 * now also use — three separate ad-hoc copies of this used to disagree on
 * which categories counted).
 */
export function isEmdExempt(workName: string): boolean {
  return isReservedWork(workName)
}

/** Indian financial year for a date (1 April boundary): 2026-07-27 -> "2026-27". */
export function indianFinancialYear(d = new Date()): string {
  const y = d.getFullYear()
  const startY = d.getMonth() >= 3 ? y : y - 1
  return `${startY}-${String((startY + 1) % 100).padStart(2, '0')}`
}

/** "18%", " -5 " -> 18, -5. Blank/unparseable -> undefined. */
export function parsePct(v: string | undefined): number | undefined {
  const s = String(v ?? '').replace(/[%,\s]/g, '')
  if (s === '') return undefined
  const n = Number(s)
  return Number.isFinite(n) ? n : undefined
}

/** Plain 2-decimal figure, no digit grouping — matches the portal's "400839.00". */
function money2(n: number | null | undefined): string {
  return n == null ? '' : n.toFixed(2)
}

/** Placeholder labels that all mean the one "price bid opened" date (kept in sync). */
const PRICE_BID_DATE_LABELS = new Set([
  'pricebidopen',
  'price bid open',
  'price bid date',
  'price bid opening date',
  'price bid opened date'
])

/**
 * Resolves one Intimation placeholder's value from the available sources, in
 * priority order: the uploaded portal HTML notice → the uploaded evaluation /
 * L-1 selection PDF → the picked Works List row. Amounts follow the office's
 * own intimation wording exactly (plain 2-decimal ECV/Contract, rounded
 * EMD @ 1.5% and ASD, and the "(Rs. 1 ½ Rs.…)" EMD expression — with ASD
 * appended only above 25% and "Exempted" for reserved works). The two date
 * placeholders share one value (see PRICE_BID_DATE_LABELS).
 *
 * Shared by the Works-List "Give Intimation" tab (row-driven) and the Tools
 * "Intimation" tool (where the row is synthesised from the uploaded L-1 with the
 * Circle/CNO/Zone resolved from the work name).
 */
export function resolveIntimationValue(
  label: string,
  notice: IntimationNotice,
  pdf: TenderEvaluation,
  row: Record<string, string>,
  office?: IntimationOffice
): string {
  const est = computeWorkAmounts(row)
  const ecv = notice.ecvRupees ?? pdf.ecvRupees ?? est.ecv ?? null
  const tenderPct = pdf.tenderPercentage ?? parsePct(tenderPercentFromRow(row))
  const contract =
    notice.contractRupees ??
    pdf.contractRupees ??
    (ecv != null && tenderPct != null ? ecv * (1 - tenderPct / 100) : null)
  // EMD @ 1.5% and ASD, rounded to the nearest rupee (6204.50/.51 -> 6205,
  // 6204.40 -> 6204) to match the office's filled samples.
  const emd = ecv != null ? Math.round(ecv * 0.015) : null
  const asd =
    ecv == null ? null : tenderPct != null && tenderPct > 25 ? Math.round((ecv * (tenderPct - 25)) / 100) : 0
  // The name of work comes from the uploaded L-1 sheet (see Give Intimation) —
  // the Works List row only supplies supporting details when its name matched.
  const workName = pdf.nameOfWork || row['Name of the work'] || ''
  const reserved = isEmdExempt(workName)

  const key = norm(label)
  // Price bid opening date = the L-1 sheet's "Server Time" (bottom-right footer),
  // i.e. when the price bid was opened — not the NIT date.
  if (PRICE_BID_DATE_LABELS.has(key)) return pdf.serverDate || pdf.noticeDate || ''

  switch (key) {
    case 'agency name':
    case 'name of the agency':
      return notice.agencyName ?? pdf.l1AgencyName ?? row['Name of the Agency'] ?? ''
    case 'address of the agency':
      return wrapAgencyAddress(notice.address ?? row['Address of the agency'] ?? '')
    case 'agency phone number':
    case 'phone number of the agency':
      return row['Phone number of the agency'] ?? ''
    case 'circle':
      return row['Circle'] || office?.circle || ''
    case 'cno':
      return row['CNO'] || office?.circleNumber || ''
    case 'zone':
      return row['Zone'] || office?.zone || ''
    case 'zoneabbr':
      return zoneAbbr(row['Zone'] || office?.zone || '')
    case 'corporation':
      return row['Corporation'] || office?.corporation || ''
    case 'corp full caps':
      return (row['Corp Full'] || office?.corporationFullName || '').toUpperCase()
    case 'financial year':
      return indianFinancialYear()
    case 'name of the work':
      return workName
    case 'nit no':
    case 'tender notice no':
      return notice.nitNo ?? pdf.noticeNo ?? row['Tender Notice No'] ?? ''
    case 'tender id':
      return pdf.tenderId ?? row['Tender ID'] ?? ''
    case 'estimate amount': {
      const raw = (row['Amount of estimate'] ?? '').replace(/,/g, '').trim()
      const n = Number(raw)
      return raw && Number.isFinite(n) ? `Rs.${n.toFixed(2)} Lakhs` : ''
    }
    case 'tender pencentage':
    case 'tender percentage':
      return tenderPct != null ? String(tenderPct) : ''
    case 'ecv':
      return money2(ecv)
    case 'contract amount':
      return money2(contract)
    case 'emd':
    case 'emd 1.5%':
      // EMD @ 1.5% is exempted for reserved works, but the ASD (charged when the
      // quote is >25% below) is independent of that exemption — keep it here so a
      // reserved work quoted >25% below still shows its ASD, not just "Exempted".
      if (reserved)
        return asd != null && asd > 0 ? `Rs. 1 ½ Rs.Exempted,ASD Rs.${asd}/-` : 'Rs. 1 ½ Rs.Exempted/-'
      if (emd == null) return ''
      return asd != null && asd > 0 ? `Rs. 1 ½ Rs.${emd},ASD Rs.${asd}/-` : `Rs. 1 ½ Rs.${emd}/-`
    case 'emd 1%':
      return ecv != null ? String(Math.round(ecv * 0.01)) : ''
    case 'asd':
      return asd != null && asd > 0 ? `ASD Rs.${asd}/-` : ''
    default:
      return ''
  }
}

/**
 * Settings' Document Templates section preview: fills every Intimation
 * template variant's {{Placeholder}}s straight off the shared
 * WorkOrderAgreementFields sample data (same fields Work Order/Agreement
 * preview tiles use), rather than the real notice/PDF/Works-List-row sources
 * resolveIntimationValue reads from — there's no upload to preview with, just
 * one office's sample values. Keys match the literal placeholder text in
 * resources/intimation-template*.docx exactly (case-sensitive).
 */
export function intimationPlaceholders(f: WorkOrderAgreementFields): Record<string, string> {
  const ecv = Number(f.ecvRupees)
  const tenderPct = Number(f.tenderPercent)
  const contract = Number(f.contractRupees)
  const estLakhs = Number(f.estimateLakhs)
  const hasEcv = f.ecvRupees.trim() !== '' && Number.isFinite(ecv)
  const hasPct = f.tenderPercent.trim() !== '' && Number.isFinite(tenderPct)
  const emd = hasEcv ? Math.round(ecv * 0.015) : null
  const asd = !hasEcv ? null : hasPct && tenderPct > 25 ? Math.round((ecv * (tenderPct - 25)) / 100) : 0
  const reserved = isEmdExempt(f.nameOfWork)
  const emdText = reserved
    ? asd && asd > 0
      ? `Rs. 1 ½ Rs.Exempted,ASD Rs.${asd}/-`
      : 'Rs. 1 ½ Rs.Exempted/-'
    : emd == null
      ? ''
      : asd && asd > 0
        ? `Rs. 1 ½ Rs.${emd},ASD Rs.${asd}/-`
        : `Rs. 1 ½ Rs.${emd}/-`
  const priceBidDate = f.noticeDate || f.intimationDate
  return {
    'Address of the agency': wrapAgencyAddress(f.address),
    'Agency Name': f.agencyName,
    CNO: f.cno,
    Circle: f.circle,
    'Contract Amount': Number.isFinite(contract) ? money2(contract) : '',
    'Corp Full Caps': (f.corporationFullName ?? '').toUpperCase(),
    Corporation: f.corporation,
    ECV: hasEcv ? money2(ecv) : '',
    'EMD 1.5%': emdText,
    'Estimate Amount': Number.isFinite(estLakhs) ? `Rs.${estLakhs.toFixed(2)} Lakhs` : '',
    'Financial year': f.financialYear,
    'Name of the work': f.nameOfWork,
    'Nit No': f.noticeNo,
    'Price Bid opening date': priceBidDate,
    Pricebidopen: priceBidDate,
    'Tender Id': f.tenderId,
    'Tender Pencentage': hasPct ? String(tenderPct) : '',
    Zone: f.zone,
    ZoneAbbr: zoneAbbr(f.zone),
    'agency phone number': f.phone
  }
}

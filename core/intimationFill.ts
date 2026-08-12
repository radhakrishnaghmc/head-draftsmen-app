import { computeWorkAmounts, tenderPercentFromRow } from './worksAmounts'
import { wrapAgencyAddress } from './workOrderAgreement'
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
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

/**
 * EMD @ 1.5% is exempted for works reserved for a particular category — the
 * work name carries a "reserved for SC/ST/WLCCS/…" tag. Matches any
 * "reserved for <category>" wording, case-insensitively.
 */
export function isEmdExempt(workName: string): boolean {
  return /reserved\s+for\b/i.test(workName)
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
 * own intimation wording exactly (plain 2-decimal ECV/Contract, floored
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
  // EMD @ 1.5% and ASD, floored to match the office's filled samples.
  const emd = ecv != null ? Math.floor(ecv * 0.015) : null
  const asd =
    ecv == null ? null : tenderPct != null && tenderPct > 25 ? Math.floor((ecv * (tenderPct - 25)) / 100) : 0
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
      return ecv != null ? String(Math.floor(ecv * 0.01)) : ''
    case 'asd':
      return asd != null && asd > 0 ? `ASD Rs.${asd}/-` : ''
    default:
      return ''
  }
}

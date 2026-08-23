/**
 * "Notice for Non-conclusion of Agreement" — issued to an L-1 bidder who won
 * a tender (LOA issued) but hasn't come in to conclude the agreement within
 * the stipulated time. Same family as the SE LOA intimation (core/loaSe.ts) —
 * reuses its Indian-amount/words formatting so the two documents read
 * consistently, since this Notice is literally the LOA's own follow-up
 * letter. Issued from either a Zone-level (SE) office or a Circle-level (EE)
 * office — two separate bundled templates (issue-notice-template.docx /
 * issue-notice-ee-template.docx) carry each office's own letterhead,
 * numbering and signature, but this one function fills both: it emits every
 * placeholder either template uses (Zone/Zone Abbr for the SE format,
 * Circle/CNO for the EE one), so the caller just picks which template to
 * load and this never has to know which.
 *
 * Standalone, like the Tools-tab Intimation (core/intimationFill.ts /
 * IntimationToolTab): it works from just the uploaded L-1 selection form and
 * Online Intimation — no Works List row and no matching office login needed.
 */
import { formatIndianAmount, amountInWords, zoneAbbr, financialYearFromDate } from './loaSe'
import { extractItemNo, stripItemNoTag } from './bidDocument'
import type { IntimationNotice } from './intimationNotice'
import type { TenderEvaluation } from './tenderEvaluationPdf'

/** The office block this Notice's letterhead/signature names — a Zone (SE office) or a Circle+CNO (EE office). */
export interface IssueNoticeOffice {
  zone?: string
  circle?: string
  cno?: string
}

/**
 * Fields neither upload can supply — typed in by whoever is issuing the
 * Notice. The LOA No/Date are the department's own record of the earlier
 * Letter of Acceptance (not printed on the L-1 sheet or the Online
 * Intimation); the Notice Date and agency phone are likewise always
 * hand-entered (the phone is only sometimes present on the "To" block — a
 * firm vs. an individual bidder — so it's never reliably parseable).
 */
export interface IssueNoticeManualFields {
  loaNo: string
  /** "dd.mm.yyyy". */
  loaDate: string
  /** "dd.mm.yyyy" — this Notice's own issue date. */
  noticeDate: string
  agencyPhone: string
}

export const EMPTY_ISSUE_NOTICE_MANUAL_FIELDS: IssueNoticeManualFields = {
  loaNo: '',
  loaDate: '',
  noticeDate: '',
  agencyPhone: ''
}

/**
 * Builds the {{Placeholder}} → value map for the bundled Notice formats
 * (resources/issue-notice-template.docx for an SE office,
 * issue-notice-ee-template.docx for an EE office). Sources, in priority
 * order (mirrors Give Intimation / Work Order): the uploaded Online
 * Intimation → the L-1 evaluation / selection PDF.
 */
export function issueNoticePlaceholders(
  notice: IntimationNotice,
  pdf: TenderEvaluation,
  manual: IssueNoticeManualFields,
  office: IssueNoticeOffice
): Record<string, string> {
  const ecv = notice.ecvRupees ?? pdf.ecvRupees ?? null
  const tenderPct = pdf.tenderPercentage ?? null
  const contract =
    notice.contractRupees ??
    pdf.contractRupees ??
    (ecv != null && tenderPct != null ? ecv * (1 - tenderPct / 100) : null)
  const workName = stripItemNoTag(pdf.nameOfWork || '')
  const itemNo = extractItemNo(pdf.nameOfWork || '') ?? ''

  return {
    Zone: office.zone ?? '',
    'Zone Abbr': zoneAbbr(office.zone),
    Circle: office.circle ?? '',
    CNO: office.cno ?? '',
    'Agency Name': notice.agencyName ?? pdf.l1AgencyName ?? '',
    'Address of the agency': notice.address ?? '',
    'agency phone number': manual.agencyPhone,
    'Financial year': financialYearFromDate(manual.noticeDate),
    'Notice Date': manual.noticeDate,
    'Name of the work': workName,
    'Nit No': notice.nitNo ?? pdf.noticeNo ?? '',
    'Nit Date': notice.nitDate ?? pdf.noticeDate ?? '',
    'Item No': itemNo,
    'LOA No': manual.loaNo,
    'LOA Date': manual.loaDate,
    'Tender Contract Value': formatIndianAmount(contract, 2),
    'Tender Contract Value In Words': amountInWords(contract),
    'Tender Percentage': tenderPct != null ? String(tenderPct) : '',
    'Estimate Contract Value': formatIndianAmount(ecv, 2)
  }
}

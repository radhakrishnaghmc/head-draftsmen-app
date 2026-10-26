import { useEffect, useLayoutEffect, useMemo, useRef, useState, type RefObject } from 'react'
import { createPortal } from 'react-dom'
import { renderAsync } from '../lazyDocxPreview'
import { api } from '../ipc'
import {
  parseIntimationNotice,
  parseIntimationNoticeText,
  parseSupplierDetailsText,
  type IntimationNotice
} from '@core/intimationNotice'
import { parseTenderEvaluation, parseAllBidders, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { detectTenderId } from '@core/tenderAgents'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import { computeWorkAmounts, tenderPercentFromRow, corpusFundFromEcv } from '@core/worksAmounts'
import { wrapAgencyAddress } from '@core/workOrderAgreement'
import { extractItemNo, stripItemNoTag } from '@core/bidDocument'
import { DROP_PARAGRAPH_MARKER } from '@core/dropMarker'
import {
  zoneAbbr,
  financialYearFromDate,
  formatIndianAmount,
  amountInWords
} from '@core/loaSe'
import { resolveIntimationValue } from '@core/intimationFill'
import { type NoteBidder, summarizeNonResponsiveness } from '@core/noteSubmitted'
import {
  buildBidEvaluationHtml,
  buildAgencyApprovalHtml,
  type BidEvaluationData,
  type AgencyApprovalData
} from '@core/seEvaluationNotes'
import { parseCementSteelRateLines } from '@core/cementSteelRate'
import { CEMENT_STEEL_MONTHS, cementSteelRatePeriodKey } from '@core/cementSteelRates'
import { resolveFromDirectory, entriesOf, corporationByName, CORPORATIONS } from '../zoneCircleDirectory'
import { officeScopedKey, TEMPLATE_KEYS, type Office } from '../office'
import type { PlaceholderMatch } from '@core/createDocument'
import { pdfToTextLinesOrOcr, pdfToPositionedLines } from '../pdfToText'
import { pdfPagesToDataUrls, pdfPagesToDataUrlsFromData } from '../pdfToImages'
import { base64ToUint8, PAGE_WIDTH, renderDocPreview, DOCX_PREVIEW_OPTIONS, normalizeDocxTextboxes } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell, IconCheck } from './Icons'
import type { ExcelTable } from '@core/types'
import type { AgreementBundleFile, VerifyDocItem } from '../../electron/ipc-contract'
import { VerifyButton } from './VerifyButton'
import PageMarginGuides from './PageMarginGuides'
import { applyPageMargins, type PageMargins } from '@core/docMargins'
import DocFieldsEditor, { groupLabels, humanizeLabel, type DocField, type DocFieldGroup } from './DocFieldsEditor'
import { variantForCorporation, corporationDefaultVariant, INTIMATION_TEMPLATE_VARIANTS } from '@core/workOrderTemplateVariants'

interface Props {
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
  /** The chosen office. A Zone with no Circle = Superintending Engineer (zonal) office → the LOA format. */
  office: Office
  /**
   * DOM node (rendered by the page header, App.tsx) to portal the "Download
   * all documents" button into — puts it in the page-head-action slot next
   * to the title, matching the Agreement page, instead of buried in the
   * body below the upload buttons.
   */
  headerActionRef?: RefObject<HTMLDivElement | null>
  /**
   * Whether this tab is the one currently on screen. App.tsx keeps this
   * component mounted (but hidden) when the user switches to another tab
   * (KeepAlive), so re-reading the office's chosen Intimation template
   * variant only on mount would keep showing whatever variant was picked
   * when the user first opened this tab — even after they change it in
   * Settings and switch back. Toggling true re-triggers that read.
   */
  active?: boolean
  /**
   * Fired right after the Intimation/LOA is actually issued (downloaded or
   * printed, main doc or the "Download all" bundle) — App.tsx turns this into
   * a pre-filled Issue Notices tile, so the office never has to re-upload the
   * same Online Intimation/L1 sheet to prepare the Non-conclusion Notice.
   * Omits id/officeKey/done/deletedAt/manual (App.tsx owns those — it also
   * dedupes against an existing tile for the same work instead of adding a
   * second one).
   */
  onIntimationIssued?: (snapshot: IntimationIssuedSnapshot) => void
}

/** See `onIntimationIssued` above. */
export interface IntimationIssuedSnapshot {
  seMode: boolean
  zone?: string
  circle?: string
  cno?: string
  corporation?: string
  agencyName?: string
  address?: string
  workName?: string
  nitNo?: string
  nitDate?: string
  itemNo?: string
  ecvRupees?: number
  contractRupees?: number
  tenderPercentage?: number
  /** Seeds the Notice's "LOA Date" field — the same value already shown as this Intimation/LOA's own date (pdf.serverDate, falling back to the hand-typed one). */
  loaDateSeed?: string
}

/** The Intimation letter (or SE LOA) plus its 4 SE companion notes — each shown as a tile that expands into the preview modal, same layout as the Agreement page's document tiles. */
type ExpandedDoc = 'intimation' | 'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval'

/** Fields the SE LOA (and its 4 companion notes) need that no source (Works List / notice / L-1) can supply — the user types them. */
interface LoaManualFields {
  adminSanction: string
  period: string
  itemNo: string
  loaDate: string
  /** Which authority approved Administrative Sanction — varies per work; the TS Note / Eligibility Criteria / Bid Evaluation / Agency Approval notes all cite it. */
  asAuthority: 'zonal' | 'commissioner' | 'resolution'
  /** Only used when asAuthority is 'commissioner' — the zonal form's Ref line carries no date in the office's own samples. */
  asDate: string
  /** SSR (Schedule of Rates) year the TS Note cites, e.g. "2026-27" — defaults to the current financial year. */
  ssrYear: string
  tsNo: string
  tsDate: string
  techBidOpenDate: string
  bidEvalApprovedDate: string
  /** Cement & Steel rate circular (OCR-read from the uploaded PDF, then editable — see uploadCementSteelRates). */
  cementRate: string
  steelRate: string
  msFlatsRate: string
  memoRef: string
  memoDate: string
  rateMonth: string
  /** Eligibility Criteria — clauses that vary per work; defaults match the office's usual wording. */
  classOfRegistration: string
  emdValidity: string
  solvencyValidity: string
  graduateEngineers: string
  diplomaEngineers: string
  equipment1: string
  equipment2: string
  equipment3: string
  goMsNo: string
}

const LOA_MANUAL_DEFAULTS: LoaManualFields = {
  adminSanction: '',
  period: '3 Months',
  itemNo: '',
  loaDate: '',
  asAuthority: 'commissioner',
  asDate: '',
  ssrYear: financialYearFromDate(undefined),
  tsNo: '',
  tsDate: '',
  techBidOpenDate: '',
  bidEvalApprovedDate: '',
  cementRate: '',
  steelRate: '',
  msFlatsRate: '',
  memoRef: '',
  memoDate: '',
  rateMonth: '',
  classOfRegistration: 'Appropriate class of registration as per existing G.O.s in Telangana State.',
  emdValidity: '6 Months',
  solvencyValidity: '1 year',
  graduateEngineers: '1 No.',
  diplomaEngineers: '2 Nos.',
  equipment1: 'Tipper',
  equipment2: 'Hydraulic Excavator',
  equipment3: 'Water tanker',
  goMsNo: '66'
}

/** A staffing count typed as 0 ("0", "0 No.", "Nil") means the post isn't required — its line is left out of the document. */
function isZeroCount(v: string): boolean {
  return /^\s*(0+(\s|\.|$)|nil\b|none\b)/i.test(v)
}

/** "<Circle> Circle-<CNO>" from the Works List row, falling back to parsing it out of the name of the work (a Zone login issuing a work from an unmatched circle still needs this — same directory lookup the Work Order/Agreement page uses). */
function circleLine(row: Record<string, string>, workName: string, office: Office): string {
  let circle = (row['Circle'] ?? '').trim()
  let cno = (row['CNO'] ?? row['Circle number'] ?? '').trim()
  if (!circle) {
    const resolved = resolveFromDirectory(workName, entriesOf(office.corporation))
    circle = resolved.circle ?? ''
    cno = resolved.cno ?? ''
  }
  if (!circle) return ''
  return cno ? `${circle} Circle-${cno}` : `${circle} Circle`
}

/** "Zonal Commissioner, QBZ, CMC" or "Commissioner, CMC" — the AS-approving authority, per the manual toggle. */
function asAuthorityText(zoneAbbrCode: string, m: LoaManualFields, corp = 'CMC'): string {
  if (m.asAuthority === 'resolution') return `Standing Committee, ${corp}`
  return m.asAuthority === 'zonal' ? `Zonal Commissioner, ${zoneAbbrCode}, ${corp}` : `Commissioner, ${corp}`
}

/** The Ref-1 sentence tail: the authority, plus ", dated:DD.MM.YYYY" only for the Commissioner form — the zonal form's Ref line carries no date in the office's own samples. */
function asRefLineText(zoneAbbrCode: string, m: LoaManualFields, corp = 'CMC'): string {
  const authority = asAuthorityText(zoneAbbrCode, m, corp)
  return m.asAuthority === 'zonal' ? authority : `${authority}, dated:${m.asDate.trim() || '__________'}`
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

// A stable empty-row reference for `detailsRow` below — a fresh `{}` literal
// on every render would break the memoization of everything downstream
// (seDocValues/bidEvalData/agencyApprovalData and the SE companion notes'
// live-preview effect), re-filling and re-rendering all 4 SE documents on
// every render instead of only when the data actually changes. Same root
// cause as the Agreement/Work Order tab's selectedRow bug.
const EMPTY_ROW: Record<string, string> = {}

/** "18%", " -5 " -> 18, -5. Blank/unparseable -> undefined. */
function parsePct(v: string | undefined): number | undefined {
  const s = String(v ?? '').replace(/[%,\s]/g, '')
  if (s === '') return undefined
  const n = Number(s)
  return Number.isFinite(n) ? n : undefined
}

const NOT_RESERVED = /^(no|none|general|open|nil|n\/?a|-|not\s+reserved)$/i

/**
 * Works out a work's clean name (for the Sub / acceptance paragraph), its
 * "(Reserved for X only)" tag, and whether it's reserved. The reserved category
 * comes from the Works List "Reservation" column when set, otherwise from a
 * "(Reserved for …)" tag embedded in the name. The name is read from the row
 * with a fallback to the L-1 PDF's work name (so an empty row cell still fills),
 * always stripped of any inline reserved tag.
 */
function reservedInfo(
  row: Record<string, string>,
  pdf: TenderEvaluation
): { name: string; tag: string; isReserved: boolean } {
  // The name of work comes from the uploaded L-1 sheet ONLY (the specific work
  // being issued — the agency/amounts come from the uploads too). The Works List
  // is only used to fill supporting details (Circle, etc.) when the L-1's work
  // name matches a row; it is never the source of the name shown here.
  const rawName = (pdf.nameOfWork || '').trim()
  // Clean name + any tag embedded in the name text.
  const paren = /\(\s*reserved\s+for\s+(.+?)\s*\)/i.exec(rawName)
  const bare = paren ? null : /reserved\s+for\s+([A-Za-z/&]+)/i.exec(rawName)
  const name = paren || bare
    ? rawName
        .replace(/\(?\s*reserved\s+for[^)]*\)?/i, '')
        .replace(/\s{2,}/g, ' ')
        .replace(/\s+([,).])/g, '$1')
        .trim()
    : rawName
  // Category: prefer the Reservation column, else the tag parsed from the name.
  const col = (row['Reservation'] ?? '').trim()
  let cat = col && !NOT_RESERVED.test(col)
    ? col.replace(/reserved\s*(for)?/i, '').replace(/\bonly\b/i, '').replace(/[()]/g, '').trim()
    : ''
  if (!cat) cat = (paren?.[1] ?? bare?.[1] ?? '').replace(/\bonly\b/i, '').trim()
  // The item number (SE multi-item NITs tag it onto the name, e.g. "...(Item
  // No.3)") has its own field elsewhere — never print it as part of the name.
  return { name: stripItemNoTag(name), tag: cat ? `(Reserved for ${cat} only)` : '', isReserved: !!cat }
}

/** Whether a work is SC/ST-reserved (picks the LOA variant that omits the EMD line). */
function isReservedWork(row: Record<string, string>, pdf: TenderEvaluation): boolean {
  return reservedInfo(row, pdf).isReserved
}

/**
 * Resolves one Superintending-Engineer LOA placeholder. Shares the amount
 * sources with {@link resolveValue} (ECV / contract / tender %), formats money
 * the SE way (Indian grouping), and derives the LOA-only figures: the balance
 * EMD clause (1.5% of ECV, plus ASD = (tender%−25)% when the quote is >25%
 * below), E-Corpus @ 0.04% of ECV, the amount in words, the zone code, the
 * "(Reserved for …)" tag and the Copy-to circle line. The three fields no
 * source can supply (Admin Sanction Value, Period, Item No) come from `manual`.
 */
function resolveLoaValue(
  label: string,
  notice: IntimationNotice,
  pdf: TenderEvaluation,
  row: Record<string, string>,
  office: Office,
  manual: LoaManualFields
): string {
  const est = computeWorkAmounts(row)
  const ecv = notice.ecvRupees ?? pdf.ecvRupees ?? est.ecv ?? null
  const tenderPct = pdf.tenderPercentage ?? parsePct(tenderPercentFromRow(row))
  const contract =
    notice.contractRupees ??
    pdf.contractRupees ??
    (ecv != null && tenderPct != null ? ecv * (1 - tenderPct / 100) : null)
  const work = reservedInfo(row, pdf)

  const key = norm(label)
  switch (key) {
    case 'zone abbr':
      return zoneAbbr(office.zone)
    case 'zone':
      return office.zone ?? row['Zone'] ?? ''
    case 'corporation':
      return office.corporation ?? row['Corporation'] ?? ''
    case 'corp full caps':
      return (corporationByName(office.corporation)?.fullName ?? '').toUpperCase()
    case 'financial year':
      return financialYearFromDate(pdf.noticeDate)
    case 'loa date':
      // Date of issuance = the L-1 sheet's "Server Time" (bottom-right footer),
      // falling back to a hand-entered date only when the sheet has none.
      return pdf.serverDate || manual.loaDate
    case 'agency name':
      return notice.agencyName ?? pdf.l1AgencyName ?? row['Name of the Agency'] ?? ''
    case 'address of the agency':
      return wrapAgencyAddress(notice.address ?? row['Address of the agency'] ?? '')
    case 'agency phone number':
      return notice.phone ?? row['Phone number of the agency'] ?? ''
    case 'item no':
      return manual.itemNo
    case 'intimation lr no': {
      // "<Item No>/SE/<Zone Abbr>/<Corp>/<FY>" (MMC: "SE(M)" and "<Zone Abbr> ZONE"). One editable
      // field, so the whole number can be retyped when the office's numbering differs.
      const corp = office.corporation ?? row['Corporation'] ?? ''
      const abbr = zoneAbbr(office.zone)
      const fy = financialYearFromDate(pdf.noticeDate)
      return corp === 'MMC' ? `${manual.itemNo}/SE(M)/${abbr} ZONE/${corp}/${fy}` : `${manual.itemNo}/SE/${abbr}/${corp}/${fy}`
    }
    case 'name of the work':
      return work.name
    case 'reserved tag':
      return work.tag
    case 'circle line':
      return circleLine(row, work.name, office)
    case 'nit no':
      return notice.nitNo ?? pdf.noticeNo ?? row['Tender Notice No'] ?? ''
    case 'nit date':
      return notice.nitDate ?? pdf.noticeDate ?? ''
    case 'tender id':
      return pdf.tenderId ?? ''
    case 'price bid opening date':
      // The "Server Time" from the L-1 sheet's bottom-right footer (when the
      // price bid was opened), not the NIT date.
      return pdf.serverDate || pdf.noticeDate || ''
    case 'admin sanction value': {
      const raw = manual.adminSanction.trim()
      if (!raw) return ''
      const n = Number(raw.replace(/[,\s₹]/g, '').replace(/rs\.?/i, ''))
      return Number.isFinite(n) && /\d/.test(raw) ? formatIndianAmount(n, 2) : raw
    }
    case 'ecv':
      return formatIndianAmount(ecv, 2)
    case 'tender percentage':
      return tenderPct != null ? String(tenderPct) : ''
    case 'contract amount':
      return formatIndianAmount(contract, 2)
    case 'period of completion':
      return manual.period.trim() || (row['Completion Period'] ?? '').trim()
    case 'contract in words':
      return amountInWords(contract)
    case 'emd clause': {
      if (ecv == null) return ''
      const emd = formatIndianAmount(Math.round(ecv * 0.015), 0)
      const asd = tenderPct != null && tenderPct > 25 ? Math.round((ecv * (tenderPct - 25)) / 100) : 0
      return asd > 0 ? `Rs. ${emd}/- & ASD amount of Rs.${formatIndianAmount(asd, 0)}/-` : `Rs. ${emd}/-`
    }
    case 'e-corpus':
      return ecv != null ? formatIndianAmount(corpusFundFromEcv(ecv), 0) : ''
    case 'period number': {
      const n = (manual.period.trim() || (row['Completion Period'] ?? '').trim()).match(/\d+/)?.[0]
      return n ?? ''
    }
    case 'bg validity months': {
      const period = Number((manual.period.trim() || (row['Completion Period'] ?? '').trim()).match(/\d+/)?.[0] ?? '')
      return Number.isFinite(period) && period > 0 ? String(period + 24) : ''
    }
    case 'mmc emd clause': {
      if (ecv == null) return ''
      const emd = formatIndianAmount(Math.round(ecv * 0.015), 0)
      const asd = tenderPct != null && tenderPct > 25 ? Math.round((ecv * (tenderPct - 25)) / 100) : 0
      return asd > 0 ? `of Rs.${emd}/- & ASD of Rs.${formatIndianAmount(asd, 0)}/-` : `Rs. ${emd}/-`
    }
    case 'mmc reserved emd clause': {
      if (ecv == null) return ''
      const emd = formatIndianAmount(Math.round(ecv * 0.025), 0)
      const base = `EMD Exempted. However, an amount of Rs.${emd}/- towards 2½% EMD shall be recovered from respective RA bills as stipulated under G.O.Ms.No.59, I&CAD (Reforms) Department, Dt:21-05-2018.`
      const asd = tenderPct != null && tenderPct > 25 ? Math.round((ecv * (tenderPct - 25)) / 100) : 0
      return asd > 0
        ? `${base} & Submitted ASD of Rs.${formatIndianAmount(asd, 0)}/- through online / Irrevocable Bank Guarantee from any nationalized / scheduled banks in favour of the Commissioner, ${office.corporation ?? row['Corporation'] ?? ''}.`
        : base
    }
    default:
      return ''
  }
}

/**
 * Resolves one placeholder shared by the TS Note and Eligibility Criteria
 * templates (SE-office, issued alongside the LOA) — a single switch since the
 * two templates' placeholder sets overlap heavily; each template's fill only
 * consumes the labels it actually contains, same convention as
 * zonalDocsPlaceholders in core/workOrderAgreement.ts.
 */
function resolveTsNoteValue(
  label: string,
  pdf: TenderEvaluation,
  row: Record<string, string>,
  office: Office,
  manual: LoaManualFields
): string {
  const work = reservedInfo(row, pdf)
  const zoneAbbrCode = zoneAbbr(office.zone)
  const est = computeWorkAmounts(row)
  const ecv = pdf.ecvRupees ?? est.ecv ?? null
  // No Works List row (or a blank "Amount of estimate" cell) shouldn't leave
  // the SE companion notes' Estimate Cost printing blank when the uploaded L-1
  // sheet's own Estimated Contract Value is right there — same fallback
  // deriveFields/standaloneRowFromSources use for the Agreement/Work Order.
  const rowEstimateLakhs = (row['Amount of estimate'] ?? '').replace(/,/g, '').trim()
  const estimateLakhs = rowEstimateLakhs || (ecv != null ? (ecv / 100000).toFixed(2) : '')

  const key = norm(label)
  switch (key) {
    case 'zone':
      return office.zone ?? ''
    case 'zone abbr':
      return zoneAbbrCode
    case 'corporation':
      return office.corporation ?? ''
    case 'financial year':
      return financialYearFromDate(pdf.noticeDate)
    case 'ssr year':
      return manual.ssrYear.trim() || financialYearFromDate(undefined)
    case 'name of the work':
      return work.name
    case 'item no':
      return manual.itemNo
    case 'ee circle':
      return circleLine(row, work.name, office)
    case 'estimate lakhs':
      return estimateLakhs
    case 'ecv':
      return ecv != null ? formatIndianAmount(ecv, 2) : ''
    case 'ecv lakhs':
      return ecv != null ? (ecv / 100000).toFixed(2) : ''
    case 'emd 1%':
      return ecv != null ? formatIndianAmount(Math.round(ecv * 0.01), 0) : ''
    case 'period of completion':
      return manual.period.trim()
    case 'as authority':
      return asAuthorityText(zoneAbbrCode, manual, (office.corporation ?? '').trim() || 'CMC')
    case 'as ref line':
      return asRefLineText(zoneAbbrCode, manual, (office.corporation ?? '').trim() || 'CMC')
    case 'ts no':
      return manual.tsNo.trim()
    case 'ts date':
      return manual.tsDate.trim()
    case 'cement rate':
      return manual.cementRate.trim()
    case 'steel rate':
      return manual.steelRate.trim()
    case 'ms flats rate':
      return manual.msFlatsRate.trim()
    case 'memo ref':
      return manual.memoRef.trim()
    case 'memo date':
      return manual.memoDate.trim()
    case 'rate month':
      return manual.rateMonth.trim()
    case 'class of registration':
      return manual.classOfRegistration.trim()
    case 'emd validity':
      return manual.emdValidity.trim()
    case 'solvency validity':
      return manual.solvencyValidity.trim()
    case 'graduate engineers':
      return isZeroCount(manual.graduateEngineers) ? DROP_PARAGRAPH_MARKER : manual.graduateEngineers.trim()
    case 'diploma engineers':
      return isZeroCount(manual.diplomaEngineers) ? DROP_PARAGRAPH_MARKER : manual.diplomaEngineers.trim()
    case 'equipment 1':
      return manual.equipment1.trim()
    case 'equipment 2':
      return manual.equipment2.trim()
    case 'equipment 3':
      return manual.equipment3.trim()
    case 'go ms no':
      return manual.goMsNo.trim()
    default:
      return ''
  }
}

/** Assembles the Sub/Ref fields shared by the Bid Evaluation and Agency Approval notes (see core/seEvaluationNotes.ts). */
function seRefsFor(
  pdf: TenderEvaluation,
  notice: IntimationNotice,
  row: Record<string, string>,
  office: Office,
  manual: LoaManualFields
) {
  const work = reservedInfo(row, pdf)
  const zoneAbbrCode = zoneAbbr(office.zone)
  const corp = (office.corporation ?? row['Corporation'] ?? '').trim() || 'CMC'
  const ecv = pdf.ecvRupees ?? computeWorkAmounts(row).ecv ?? null
  const rowEstimateLakhs = (row['Amount of estimate'] ?? '').replace(/,/g, '').trim()
  return {
    zoneAbbr: zoneAbbrCode,
    corp,
    workName: work.name,
    estimateLakhs: rowEstimateLakhs || (ecv != null ? (ecv / 100000).toFixed(2) : ''),
    asAuthority: asAuthorityText(zoneAbbrCode, manual, corp),
    asRefLine: asRefLineText(zoneAbbrCode, manual, corp),
    tsNo: manual.tsNo.trim(),
    tsDate: manual.tsDate.trim(),
    financialYear: financialYearFromDate(pdf.noticeDate),
    nitNo: notice_nitNo(pdf, row),
    nitDate: notice.nitDate || pdf.noticeDate || '',
    itemNo: manual.itemNo.trim(),
    // The L-1 sheet's own "Server Time" footer (bottom-right corner) is when
    // the technical bid was opened — falls back to the hand-typed field only
    // when no L-1 was uploaded.
    techBidOpenDate: pdf.serverDate || manual.techBidOpenDate.trim()
  }
}

/** NIT No for the SE evaluation notes — the L-1 sheet's own notice number, falling back to the Works List row. */
function notice_nitNo(pdf: TenderEvaluation, row: Record<string, string>): string {
  return pdf.noticeNo || row['Tender Notice No'] || ''
}

/**
 * Give Intimation — fills the bundled Intimation format (a .docx mail-merge
 * template with {{placeholders}}) for a Works List work: most fields come from
 * the picked row, while the agency address (and, when present, agency name /
 * NIT No / ECV / contract value) come from the uploaded Online Intimation
 * (portal HTML or LOA PDF) plus the L-1 selection form. Shows a live
 * docx-preview of the filled letter with Word / PDF / Print. (Note Submitted
 * lives on the Agreement & Work Order tab.)
 */
export default function GiveIntimationTab({ tables, onChange, office, headerActionRef, active, onIntimationIssued }: Props) {
  const table = tables[0] ?? null

  // Whether headerActionRef's DOM node is actually mounted yet — reading
  // headerActionRef.current directly during render is unreliable (refs
  // attach during commit, after render runs, so the very first render
  // always sees it as null and never gets another render to correct that
  // unless something else also changes). This effect re-renders once the
  // node is there, so the "Download all documents" button reliably ends up
  // in the header instead of silently staying in its (now-removed) inline
  // fallback forever.
  const [headerMounted, setHeaderMounted] = useState(false)
  useLayoutEffect(() => {
    setHeaderMounted(!!headerActionRef?.current)
  })

  // A Zone chosen with no Circle is the Superintending Engineer (zonal) office,
  // which issues the "Letter of Acceptance" format instead of the EE Intimation.
  const seMode = !!office.zone?.trim() && !office.circle?.trim()

  // `office` plus its corporation's full name (for the letterhead's {{Corp Full
  // Caps}}) — resolveIntimationValue is a core module and can't reach into the
  // renderer's zoneCircleDirectory itself, so it's resolved here.
  const intimationOffice = useMemo(
    () => ({ ...office, corporationFullName: corporationByName(office.corporation)?.fullName }),
    [office]
  )

  const [templateB64, setTemplateB64] = useState<string | null>(null)
  const [labels, setLabels] = useState<string[]>([])
  const [loadError, setLoadError] = useState<string | null>(null)
  const [manual, setManual] = useState<LoaManualFields>(LOA_MANUAL_DEFAULTS)

  const [rowIndex, setRowIndex] = useState(0)
  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  // eGP "Supplier Details" page — supplies the agency's phone number (which
  // the Online Intimation letter never carries) plus a cleaner name/address.
  const [supplierDetailsName, setSupplierDetailsName] = useState('')
  const [pdfEval, setPdfEval] = useState<TenderEvaluation | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pdfStatus, setPdfStatus] = useState<string | null>(null)
  // Whether the uploaded L1 form's work matched a Works List row. false means
  // it matched none, so the letter would fill from an unrelated row — the
  // preview is gated until the work is added to the Works List. null = not yet
  // determined (no L1 uploaded, or no works to match against).
  const [worksRowMatched, setWorksRowMatched] = useState<boolean | null>(null)
  const [values, setValues] = useState<Record<string, string>>({})
  // Hand-typed overrides for the letter's {{placeholder}} values, keyed by
  // placeholder label — typed in the expanded preview's editor panel. A label
  // present here wins over the value resolved from the uploads/Works List row,
  // including when it's been blanked on purpose.
  const [intimationEdits, setIntimationEdits] = useState<Record<string, string>>({})

  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf' | 'bundle' | 'notice' | 'supplier' | 'nonresp'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  // Which format "Download all documents" saves every document as. Defaults
  // to Word since PDF conversion is the slow part.
  const [bundleFormat, setBundleFormat] = useState<'docx' | 'pdf'>('docx')
  // Live "Download all documents" progress (see the same field on the
  // Agreement/Work Order tab) — filling locally, then the main process
  // writing/converting, so the button doesn't sit on a static "Preparing…"
  // that reads as a frozen screen on a multi-document SE bundle.
  const [bundleProgress, setBundleProgress] = useState<{ phase: 'preparing' | 'saving'; done: number; total: number } | null>(null)
  // Which document tile is expanded into the full preview modal — mirrors the
  // Agreement page's tile-grid layout instead of showing every document's full
  // preview inline at once.
  const [expanded, setExpanded] = useState<ExpandedDoc | null>(null)
  const [expandedPages, setExpandedPages] = useState(0)
  // Page margins (mm) dragged on a document's preview; a document with no entry uses its template's own.
  // Kept between sessions, like the Agreement page's.
  const [docMargins, setDocMargins] = useState<Partial<Record<ExpandedDoc, PageMargins>>>(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('intimationDocPageMargins') ?? 'null')
      if (saved && typeof saved === 'object') return saved
    } catch {
      /* fall through to the templates' margins */
    }
    return {}
  })
  useEffect(() => {
    localStorage.setItem('intimationDocPageMargins', JSON.stringify(docMargins))
  }, [docMargins])
  function setMarginsFor(kind: ExpandedDoc, m: PageMargins | null) {
    setDocMargins((prev) => {
      const next = { ...prev }
      if (m) next[kind] = m
      else delete next[kind]
      return next
    })
  }

  // SE-only: 4 companion documents issued alongside the LOA. TS Note and
  // Eligibility Criteria are docx templates (like the LOA itself, keyed by
  // 'tsNote' / 'eligibility'); Bid Evaluation and Agency Approval are built as
  // HTML (dynamic bidder tables), same technique as the EE Note Submitted.
  const [seDocB64, setSeDocB64] = useState<Record<string, string>>({})
  const [seDocLabels, setSeDocLabels] = useState<Record<string, string[]>>({})
  // Every bidder on the uploaded L-1 selection form (not just L-1) — feeds the
  // Bid Evaluation / Agency Approval notes' tables.
  const [allBidders, setAllBidders] = useState<NoteBidder[]>([])
  const [cementSteelBusy, setCementSteelBusy] = useState(false)
  const [cementSteelError, setCementSteelError] = useState<string | null>(null)
  const [cementSteelFileName, setCementSteelFileName] = useState('')
  // Month/year the user wants the Cement & Steel Rates circular for — drives
  // fetchCementSteelRatesAuto's lookup against the department site's list.
  const [rateFetchMonth, setRateFetchMonth] = useState('')
  const [rateFetchYear, setRateFetchYear] = useState('')
  // From the uploaded "List of Bidders Made Non-Responsive" sheet — feeds the
  // Bid Evaluation note's participated/responsive/non-responsive counts, the
  // same way the EE Note Submitted's non-responsive upload works.
  const [nonRespCount, setNonRespCount] = useState(0)
  const [nonRespFileName, setNonRespFileName] = useState('')

  const noticeInputRef = useRef<HTMLInputElement>(null)
  const supplierDetailsInputRef = useRef<HTMLInputElement>(null)
  const pdfInputRef = useRef<HTMLInputElement>(null)
  const cementSteelInputRef = useRef<HTMLInputElement>(null)
  const nonRespInputRef = useRef<HTMLInputElement>(null)
  const previewRef = useRef<HTMLDivElement>(null)
  const tsNotePreviewRef = useRef<HTMLDivElement>(null)
  const eligibilityPreviewRef = useRef<HTMLDivElement>(null)
  const bidEvalPreviewRef = useRef<HTMLDivElement>(null)
  const agencyApprovalPreviewRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)
  const expandedRef = useRef<HTMLDivElement>(null)

  const selectedRow = table && table.rows.length > 0 ? table.rows[Math.min(rowIndex, table.rows.length - 1)] : null
  // The row whose details (Circle, Reservation, …) support the letter. When the
  // uploaded L-1 matched no Works List row, `selectedRow` is just row 0 (a
  // different work), so ignore it — the SE LOA fills those details blank rather
  // than borrowing another work's. A matched L-1 sets rowIndex to that row.
  // Before either upload, stay blank too — SE mode shows its document catalog
  // pre-upload (see seDocsReady) and must not show row 0's unrelated work.
  const detailsRow = !notice || !pdfEval ? EMPTY_ROW : worksRowMatched === false ? EMPTY_ROW : (selectedRow ?? EMPTY_ROW)

  // A previously-uploaded Online Intimation left over in state from a
  // different work must never lend its agency name/address to this L-1's
  // letter (or get written back into this row's Works List entry) just
  // because both happen to be in memory together — only use it once it's
  // confirmed to describe the same work as the uploaded L-1.
  const workMatch = useMemo(() => (notice && pdfEval ? checkSameWork(notice, pdfEval) : null), [notice, pdfEval])
  const workMismatch = workMatch?.status === 'mismatch'
  const effectiveNotice = workMismatch ? null : notice

  // A reserved (SC/ST) work uses the LOA variant that omits the EMD balance item.
  const seReserved = seMode && isReservedWork(detailsRow, pdfEval ?? {})

  // Load the bundled format for this office/work (SE LOA — reserved or not —
  // when zone-level, otherwise the EE Intimation) and read its placeholders.
  // Reloads if the office kind flips, the selected work's reserved status does,
  // or this tab becomes active again — App.tsx keeps this component mounted
  // across tab switches (KeepAlive), so re-checking only on mount would miss a
  // template variant changed in Settings while this tab sat hidden.
  useEffect(() => {
    if (active === false) return
    let cancelled = false
    setTemplateB64(null)
    setLabels([])
    void (async () => {
      try {
        const intimationVariant = variantForCorporation(
          localStorage.getItem(officeScopedKey(TEMPLATE_KEYS.intimation, office)) ?? corporationDefaultVariant(office.corporation, undefined),
          office.corporation,
          INTIMATION_TEMPLATE_VARIANTS
        )
        const b64 = seMode
          ? await api.loaSeTemplate(seReserved, office.corporation)
          : await api.intimationTemplate(intimationVariant)
        const found = await api.findPlaceholdersInDocument(b64)
        if (cancelled) return
        setTemplateB64(b64)
        setLabels(found)
      } catch (e) {
        if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [seMode, seReserved, office.corporation, active])

  // Load the TS Note / Eligibility Criteria docx templates once, when the
  // office is zone-level — the 2 SE companion notes that use a fixed template
  // (Bid Evaluation / Agency Approval are HTML-built instead, see below).
  useEffect(() => {
    if (!seMode) return
    let cancelled = false
    void (async () => {
      try {
        const [tsB64, eligB64] = await Promise.all([api.tsNoteTemplate(), api.eligibilityCriteriaTemplate()])
        const [tsLabels, eligLabels] = await Promise.all([
          api.findPlaceholdersInDocument(tsB64),
          api.findPlaceholdersInDocument(eligB64)
        ])
        if (cancelled) return
        setSeDocB64({ tsNote: tsB64, eligibility: eligB64 })
        setSeDocLabels({ tsNote: tsLabels, eligibility: eligLabels })
      } catch (e) {
        if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [seMode])

  // Re-fill every placeholder whenever the row / notice / PDF / manual fields change.
  useEffect(() => {
    if (labels.length === 0) return
    setValues(() => {
      const next: Record<string, string> = {}
      for (const label of labels)
        next[label] = seMode
          ? resolveLoaValue(label, effectiveNotice ?? {}, pdfEval ?? {}, detailsRow, office, manual)
          : resolveIntimationValue(label, effectiveNotice ?? {}, pdfEval ?? {}, detailsRow, intimationOffice)
      return next
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [labels, rowIndex, effectiveNotice, pdfEval, table, seMode, office, intimationOffice, manual, worksRowMatched])

  // The Online Intimation can be uploaded as either the portal "View Intimation
  // Notice" .html page or the printed Intimation / LOA .pdf — both carry the
  // agency, address, NIT No, ECV and contract value, so parse by file type.
  async function handleNoticeFile(file: File) {
    setActionError(null)
    setBusy('notice')
    try {
      const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
      const parsed = isPdf
        ? parseIntimationNoticeText(await pdfToTextLinesOrOcr(file))
        : parseIntimationNotice(await file.text())
      setNotice(parsed)
      setNoticeName(file.name)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // eGP "Supplier Details" page (bidder's own profile PDF) — the only source
  // with a phone number, so phone always comes from here. Address only comes
  // from here when no Online Intimation was uploaded to supply it already;
  // once an LOA is uploaded its address wins and this just tops up the phone.
  async function handleSupplierDetailsFile(file: File) {
    setActionError(null)
    setBusy('supplier')
    try {
      const lines = await pdfToTextLinesOrOcr(file)

      // Supplier Details must belong to the same tender as the uploaded L1
      // sheet — both carry the portal's own numeric Tender ID, so compare
      // that directly rather than the free-text NIT No / agency name.
      const supplierTenderId = detectTenderId(lines)
      if (pdfEval?.tenderId && supplierTenderId && pdfEval.tenderId !== supplierTenderId) {
        setActionError(
          `Supplier Details is for Tender ID ${supplierTenderId} but the L1 selection form is for Tender ID ${pdfEval.tenderId}. Upload the Supplier Details for the same tender.`
        )
        return
      }

      const parsed = parseSupplierDetailsText(lines)
      setNotice((prev) => {
        const base = prev ?? {}
        return {
          ...base,
          ...(parsed.agencyName && !base.agencyName ? { agencyName: parsed.agencyName } : {}),
          ...(parsed.address && !base.address ? { address: parsed.address } : {}),
          ...(parsed.phone ? { phone: parsed.phone } : {})
        }
      })
      setSupplierDetailsName(file.name)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // The L-1 selection / evaluation PDF fills the letter's tender fields and
  // updates that work's Works List row (matched by Name of Work), auto-selecting it.
  async function handlePdfFile(file: File) {
    setBusy('pdf')
    setActionError(null)
    setPdfStatus(null)
    // Clear any previously-loaded L1's data up front — if this file fails to
    // parse below, a stale pdfEval from an earlier (possibly different) work
    // must not keep being shown next to the newly-uploaded Online Intimation,
    // which would make the two halves of the letter describe different works.
    setPdfEval(null)
    setPdfName('')
    setAllBidders([])
    setWorksRowMatched(null)
    try {
      const lines = await pdfToTextLinesOrOcr(file)
      const ev = parseTenderEvaluation(lines)
      if (!ev.nameOfWork && !ev.tenderId) {
        throw new Error("Couldn't read tender details from that PDF — is it the Commercial Evaluation / Stage Selected page?")
      }
      setPdfEval(ev)
      setPdfName(file.name)
      // The sheet's own "(Item No.13)" tag (NIT line, or tagged onto the Name of Work) is
      // this work's item number — a newly uploaded L1 describes a different work, so it
      // replaces whatever was there; with no tag on the sheet, the typed value stays.
      const itemFromSheet = ev.itemNo ?? extractItemNo(ev.nameOfWork ?? '')
      if (itemFromSheet) setManual((m) => ({ ...m, itemNo: itemFromSheet }))
      // Every bidder on the sheet (not just L-1) — feeds the SE Bid Evaluation /
      // Agency Approval notes' tables.
      setAllBidders(parseAllBidders(lines))

      if (table && ev.nameOfWork) {
        let embeddings: { rowNameVectors: number[][]; evalNameVectors: number[][] } | undefined
        const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
        if (nameHeader) {
          try {
            const [rowNameVectors, evalNameVectors] = await Promise.all([
              api.embedTexts(table.rows.map((r) => r[nameHeader] ?? '')),
              api.embedTexts([ev.nameOfWork])
            ])
            embeddings = { rowNameVectors, evalNameVectors }
          } catch {
            embeddings = undefined
          }
        }
        // A leftover Online Intimation from a different work must not lend its
        // agency name/address into this row — only pass it through once it's
        // confirmed to describe the same work as the just-uploaded L1.
        const noticeForRow = notice && checkSameWork(notice, ev).status !== 'mismatch' ? notice : undefined
        // Finds the row so its supporting details fill the letter and select
        // it (an embedding/wording-drift match has no exact name to
        // re-derive from, so without this the selection would stay on row 0
        // and fill a different work), AND persists the derived fields
        // (ECV, Tender ID, Contract Amount, …) back to the Works List the
        // same way the dedicated "Update from L1" flow (WorksListL1Update)
        // already does — an L1/Intimation uploaded here is no less
        // authoritative than one uploaded there, and previously updated
        // nothing here at all, so the two flows silently disagreed.
        const { table: updated, matchedCount, matchedRowIndices } = updateWorksListFromEvaluations(
          table,
          [ev],
          embeddings,
          noticeForRow
        )
        if (matchedCount > 0) {
          const idx = matchedRowIndices[0]
          if (idx != null && idx >= 0) setRowIndex(idx)
          setWorksRowMatched(true)
          onChange(updated)
          setPdfStatus(`Matched "${ev.nameOfWork}" to a Works List row — its details fill the letter and the Works List is updated.`)
        } else {
          setWorksRowMatched(false)
          setPdfStatus(`Read the PDF. "${ev.nameOfWork}" isn't in the Works List — the letter fills from the uploaded L1 / Intimation.`)
        }
      }
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function fillTemplate(): Promise<string> {
    if (!templateB64) throw new Error('Intimation format not loaded yet.')
    const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
    // Hand edits win over the resolved values, so the tile, the expanded page,
    // Word, PDF, Print and the bundle all carry what the editor panel shows.
    const filled = await api.fillPlaceholdersInDocument(templateB64, resolved, { ...values, ...intimationEdits })
    return applyPageMargins(filled, docMargins.intimation ?? null)
  }

  // A different work (another upload / Works List row) invalidates those
  // overrides — they were typed for the previous work's letter.
  const intimationWorkKey = `${values['Name of the work'] ?? ''}|${values['NIT No'] ?? ''}`
  useEffect(() => {
    setIntimationEdits({})
  }, [intimationWorkKey])

  // The letter number and date lead the panel (they head the letter), ahead of the
  // bucketed office / work / amount groups.
  const intimationGroups = useMemo(() => {
    const top: DocField[] = []
    if (labels.includes('Intimation Lr No')) top.push({ label: 'Intimation Lr No', title: 'Intimation Lr. No', wide: true })
    if (labels.includes('LOA Date')) top.push({ label: 'LOA Date', title: 'Intimation Date' })
    const rest = groupLabels(labels.filter((l) => !top.some((f) => f.label === l)))
    return top.length > 0 ? [{ title: 'Intimation', fields: top }, ...rest] : rest
  }, [labels])

  /**
   * Snapshots everything the Non-conclusion Notice needs, straight off the
   * same sources (and the same fallback order) resolveLoaValue/
   * resolveIntimationValue just used to fill the Intimation/LOA itself — so
   * the two documents always agree. Called right after the Intimation is
   * actually issued (see the 3 call sites below). Returns null when there's
   * nothing meaningful to notice yet (no agency, no work).
   */
  function buildNoticeSnapshot(): IntimationIssuedSnapshot | null {
    const n = effectiveNotice ?? {}
    const p = pdfEval ?? {}
    const row = detailsRow
    const agencyName = n.agencyName ?? p.l1AgencyName ?? row['Name of the Agency'] ?? ''
    const workName = stripItemNoTag(p.nameOfWork || row['Name of the work'] || '')
    if (!agencyName || !workName) return null
    const ecv = n.ecvRupees ?? p.ecvRupees ?? computeWorkAmounts(row).ecv ?? undefined
    const tenderPercentage = p.tenderPercentage ?? parsePct(tenderPercentFromRow(row)) ?? undefined
    const contractRupees =
      n.contractRupees ??
      p.contractRupees ??
      (ecv != null && tenderPercentage != null ? ecv * (1 - tenderPercentage / 100) : undefined)
    return {
      seMode,
      zone: office.zone || row['Zone'] || undefined,
      circle: office.circle || row['Circle'] || undefined,
      cno: office.circleNumber || row['CNO'] || undefined,
      corporation: office.corporation || row['Corporation'] || undefined,
      agencyName,
      address: n.address ?? row['Address of the agency'] ?? undefined,
      workName,
      nitNo: n.nitNo ?? p.noticeNo ?? row['Tender Notice No'] ?? undefined,
      nitDate: n.nitDate ?? p.noticeDate ?? undefined,
      itemNo: seMode ? manual.itemNo || undefined : undefined,
      ecvRupees: ecv ?? undefined,
      contractRupees: contractRupees ?? undefined,
      tenderPercentage: tenderPercentage ?? undefined,
      loaDateSeed: p.serverDate || manual.loaDate || undefined
    }
  }

  function notifyIntimationIssued() {
    if (!onIntimationIssued) return
    const snapshot = buildNoticeSnapshot()
    if (snapshot) onIntimationIssued(snapshot)
  }

  // The uploaded L1 form's work matched no Works List row. The letter takes its
  // name of work (and amounts, agency) from the uploads themselves — not the row
  // — so a no-match never blocks the letter; the Works List is only used to fill
  // supporting details when its name matched. It's surfaced as a soft note only.
  const noWorksRowMatch = worksRowMatched === false
  // The document tile(s) always show, regardless of office (EE or SE) or
  // whether anything's been uploaded yet — blank fields fill in as uploads
  // land, same as SE's catalog already did; EE no longer needs an upload
  // just to see the Intimation tile.
  const showBody = true

  // Live docx thumbnail of the filled letter (tile preview) — refreshed
  // whenever the values change, so this deliberately stays on the fast,
  // approximate docx-preview.js render rather than the accurate
  // Chromium-rendered one: re-rendering on every keystroke through
  // LibreOffice would make typing feel sluggish. The expanded modal below
  // (opened by an explicit click) is where the accurate render belongs.
  useEffect(() => {
    if (!templateB64 || !showBody) {
      if (previewRef.current) previewRef.current.innerHTML = ''
      return
    }
    const container = previewRef.current
    if (!container) return
    setActionError(null)
    void (async () => {
      try {
        const filled = await fillTemplate()
        container.innerHTML = ''
        await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
        normalizeDocxTextboxes(container)
      } catch (e) {
        setActionError(e instanceof Error ? e.message : String(e))
      }
    })()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templateB64, values, intimationEdits, showBody])

  // ---- SE-only companion notes: TS Note, Eligibility Criteria (docx
  // templates), Bid Evaluation & Agency Approval (HTML-built, dynamic bidder
  // tables). Shown as soon as their templates load — same "always browsable"
  // catalog as the LOA above. ----
  const seDocsReady = seMode && !!seDocB64.tsNote && !!seDocB64.eligibility

  const seDocValues = useMemo(() => {
    const out: Record<string, Record<string, string>> = {}
    for (const kind of ['tsNote', 'eligibility']) {
      const kindLabels = seDocLabels[kind] ?? []
      const vals: Record<string, string> = {}
      for (const label of kindLabels) vals[label] = resolveTsNoteValue(label, pdfEval ?? {}, detailsRow, office, manual)
      out[kind] = vals
    }
    return out
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [seDocLabels, pdfEval, detailsRow, office, manual])

  const bidEvalData: BidEvaluationData = useMemo(
    () => ({
      ...seRefsFor(pdfEval ?? {}, effectiveNotice ?? {}, detailsRow, office, manual),
      bidders: allBidders,
      nonRespCount
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [pdfEval, effectiveNotice, detailsRow, office, manual, allBidders, nonRespCount]
  )
  const agencyApprovalData: AgencyApprovalData = useMemo(
    () => ({
      ...seRefsFor(pdfEval ?? {}, effectiveNotice ?? {}, detailsRow, office, manual),
      zone: office.zone ?? '',
      // The L-1 sheet's own Server Time is also when the Technical Bid
      // evaluation was approved (same sitting), same fallback as techBidOpenDate.
      bidEvalApprovedDate: pdfEval?.serverDate || manual.bidEvalApprovedDate.trim(),
      bidders: allBidders
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [pdfEval, effectiveNotice, detailsRow, office, manual, allBidders]
  )

  async function fillSeDoc(kind: 'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval'): Promise<string> {
    if (kind === 'bidEval') return applyPageMargins(await api.noteSubmittedDocx(buildBidEvaluationHtml(bidEvalData)), docMargins.bidEval ?? null)
    if (kind === 'agencyApproval')
      return applyPageMargins(await api.noteSubmittedDocx(buildAgencyApprovalHtml(agencyApprovalData)), docMargins.agencyApproval ?? null)
    const b64 = seDocB64[kind]
    if (!b64) throw new Error('Format not loaded yet.')
    const kindLabels = seDocLabels[kind] ?? []
    const resolved: PlaceholderMatch[] = kindLabels.map((label) => ({ label, column: label, score: 1 }))
    const filled = await api.fillPlaceholdersInDocument(b64, resolved, seDocValues[kind] ?? {})
    return applyPageMargins(filled, docMargins[kind] ?? null)
  }

  // The full-size, accurate (Chromium-rendered) render — used only for the
  // expanded preview modal, where a user is actually looking at one specific
  // document.
  async function renderSeDocInto(kind: 'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval', container: HTMLElement) {
    const filled = await fillSeDoc(kind)
    await renderDocPreview(base64ToUint8(filled), container)
  }

  // The small, live thumbnail shown in each of the 4 SE companion-note tiles
  // — refreshed whenever their inputs change, so this deliberately stays on
  // the fast, approximate docx-preview.js render (mirrors the main
  // Intimation tile's identical reasoning above): 4 concurrent LibreOffice
  // conversions on every keystroke would make typing feel sluggish.
  async function renderSeDocTileThumbnail(
    kind: 'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval',
    container: HTMLElement
  ) {
    const filled = await fillSeDoc(kind)
    container.innerHTML = ''
    await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
    normalizeDocxTextboxes(container)
  }

  // Live previews for the 4 SE companion notes — refreshed whenever their inputs change.
  useEffect(() => {
    if (!seDocsReady) return
    if (tsNotePreviewRef.current) void renderSeDocTileThumbnail('tsNote', tsNotePreviewRef.current).catch(() => {})
    if (eligibilityPreviewRef.current) void renderSeDocTileThumbnail('eligibility', eligibilityPreviewRef.current).catch(() => {})
    if (bidEvalPreviewRef.current) void renderSeDocTileThumbnail('bidEval', bidEvalPreviewRef.current).catch(() => {})
    if (agencyApprovalPreviewRef.current) void renderSeDocTileThumbnail('agencyApproval', agencyApprovalPreviewRef.current).catch(() => {})
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [seDocsReady, seDocValues, bidEvalData, agencyApprovalData])

  const DOC_LABEL: Record<ExpandedDoc, string> = {
    intimation: seMode ? 'Letter of Acceptance' : 'Intimation',
    tsNote: 'TS Note',
    eligibility: 'Eligibility Criteria',
    bidEval: 'Bid Evaluation Note',
    agencyApproval: 'Agency Approval Note'
  }

  function openDoc(kind: ExpandedDoc) {
    setExpanded(kind)
  }

  // Full-size preview inside the expanded modal — rendered lazily (only the
  // clicked document), same pattern as the Agreement page's document tiles.
  useEffect(() => {
    if (!expanded) return
    const container = expandedRef.current
    if (!container) return
    setActionError(null)
    // The Intimation's preview is edited in place, so hold off a moment while
    // the user is still typing — one fill + LibreOffice render per keystroke
    // would make the panel feel sluggish.
    const delay = expanded === 'intimation' ? 300 : 0
    const timer = setTimeout(() => {
      void (async () => {
        try {
          if (expanded === 'intimation') {
            const filled = await fillTemplate()
            await renderDocPreview(base64ToUint8(filled), container)
          } else {
            await renderSeDocInto(expanded, container)
          }
          setExpandedPages(container.querySelectorAll('section.docx').length)
        } catch (e) {
          setActionError(e instanceof Error ? e.message : String(e))
        }
      })()
    }, delay)
    return () => clearTimeout(timer)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [expanded, values, intimationEdits, seDocValues, bidEvalData, agencyApprovalData, docMargins])

  async function downloadSeDoc(kind: 'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval', formats: ('docx' | 'pdf')[]) {
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillSeDoc(kind)
      const names: Record<typeof kind, string> = {
        tsNote: 'TS Note',
        eligibility: 'Eligibility Criteria',
        bidEval: 'Bid Evaluation',
        agencyApproval: 'Agency Approval'
      }
      const agencyName = notice?.agencyName ?? selectedRow?.['Name of the Agency']
      const name = `${names[kind]}${agencyName ? ` - ${agencyName}` : ''}`
      const res = await api.exportCreatedDocument(filled, name, formats)
      setActionSaved(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function printSeDoc(kind: 'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval') {
    setBusy('print')
    setActionError(null)
    try {
      const filled = await fillSeDoc(kind)
      const container = printScratchRef.current
      if (!container) throw new Error('Print failed to initialize.')
      await renderDocPreview(base64ToUint8(filled), container, { forSnapshot: true })
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // Builds one VerifyDocItem per ready document (mirrors downloadAll's own
  // list-building below) — re-derives ECV/Contract/EMD/ASD from the current
  // source uploads so the Verify button can cross-check them against what's
  // actually baked into each generated document, no AI involved.
  async function getVerifyItems(): Promise<VerifyDocItem[]> {
    const items: VerifyDocItem[] = []
    const est = computeWorkAmounts(detailsRow)
    const ecv = notice?.ecvRupees ?? pdfEval?.ecvRupees ?? est.ecv ?? null
    const tenderPct = pdfEval?.tenderPercentage ?? parsePct(tenderPercentFromRow(detailsRow))
    const contract =
      notice?.contractRupees ??
      pdfEval?.contractRupees ??
      (ecv != null && tenderPct != null ? ecv * (1 - tenderPct / 100) : null)
    const reserved = isReservedWork(detailsRow, pdfEval ?? {})
    const corporation = office.corporation ? { expected: office.corporation, all: CORPORATIONS.map((c) => c.name) } : undefined
    // A reserved EE Intimation prints "Exempted" instead of a number, and a
    // reserved non-MMC SE LOA carries no EMD figure at all — skip the
    // arithmetic check in those cases rather than risk a false "error" on an
    // otherwise-correct document. Every other case (EE or SE, non-reserved,
    // and MMC's reserved 2.5%-recovered-from-RA-bills case) prints a rounded
    // EMD/ASD figure that this can cross-check.
    const amounts =
      !reserved || (seMode && office.corporation === 'MMC')
        ? { ecv, contract, tenderPct, emdPct: reserved && office.corporation === 'MMC' ? 0.025 : 0.015 }
        : undefined

    if (showBody && templateB64) {
      items.push({
        name: seMode ? 'Letter of Acceptance' : 'Intimation',
        docxBase64: await fillTemplate(),
        values,
        requiredLabels: labels.filter((l) => !/admin sanction/i.test(l)),
        amounts,
        corporation,
        reserved
      })
    }

    if (seMode && seDocsReady) {
      const names: Record<'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval', string> = {
        tsNote: 'TS Note',
        eligibility: 'Eligibility Criteria',
        bidEval: 'Bid Evaluation',
        agencyApproval: 'Agency Approval'
      }
      for (const kind of ['tsNote', 'eligibility', 'bidEval', 'agencyApproval'] as const) {
        const kindValues = kind === 'tsNote' || kind === 'eligibility' ? (seDocValues[kind] ?? {}) : {}
        items.push({
          name: names[kind],
          docxBase64: await fillSeDoc(kind),
          values: kindValues,
          requiredLabels: Object.keys(kindValues).filter((l) => !/admin sanction/i.test(l)),
          corporation,
          reserved
        })
      }
    }

    return items
  }

  // Bundle every ready document into one saved folder — same bundle API and
  // "Download all documents" bar as the Agreement page.
  async function downloadAll() {
    setBusy('bundle')
    setActionError(null)
    setActionSaved(null)
    const totalPrep = (showBody && templateB64 ? 1 : 0) + (seMode && seDocsReady ? 4 : 0)
    let prepDone = 0
    setBundleProgress({ phase: 'preparing', done: 0, total: totalPrep })
    const unsubscribe = api.onAgreementBundleProgress(({ done, total }) =>
      setBundleProgress({ phase: 'saving', done, total })
    )
    try {
      const files: AgreementBundleFile[] = []
      const agencyName = notice?.agencyName ?? selectedRow?.['Name of the Agency']
      if (showBody && templateB64) {
        files.push({
          name: `${seMode ? 'Letter of Acceptance' : 'Intimation'}${agencyName ? ` - ${agencyName}` : ''}`,
          format: bundleFormat,
          docxBase64: await fillTemplate()
        })
        prepDone += 1
        setBundleProgress({ phase: 'preparing', done: prepDone, total: totalPrep })
      }
      if (seMode && seDocsReady) {
        const names: Record<'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval', string> = {
          tsNote: 'TS Note',
          eligibility: 'Eligibility Criteria',
          bidEval: 'Bid Evaluation',
          agencyApproval: 'Agency Approval'
        }
        for (const kind of ['tsNote', 'eligibility', 'bidEval', 'agencyApproval'] as const) {
          files.push({
            name: `${names[kind]}${agencyName ? ` - ${agencyName}` : ''}`,
            format: bundleFormat,
            docxBase64: await fillSeDoc(kind)
          })
          prepDone += 1
          setBundleProgress({ phase: 'preparing', done: prepDone, total: totalPrep })
        }
      }
      if (files.length === 0) {
        setActionError('No documents are ready to download yet.')
        return
      }
      setBundleProgress({ phase: 'saving', done: 0, total: files.length })
      const res = await api.exportAgreementBundle(files)
      if (!res) {
        setActionSaved('Cancelled.')
      } else if (res.failed.length > 0) {
        setActionSaved(`Saved ${res.written.length} document(s).`)
        setActionError(
          `${res.failed.length} document(s) could not be saved: ${res.failed.join(', ')}. ` +
            `Try downloading those as Word (.docx) instead.`
        )
      } else {
        setActionSaved(`Saved ${res.written.length} document(s) to the chosen folder.`)
      }
      if (res) notifyIntimationIssued()
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      unsubscribe()
      setBundleProgress(null)
      setBusy(null)
    }
  }

  /** Shared by the manual upload and the auto-fetch below: OCRs a Cement &
   * Steel rate circular's page images and fills the editable rate fields
   * from whatever the parser finds (never overwrites a field it can't read). */
  async function fillCementSteelRatesFromPages(dataUrls: string[]) {
    const sheet = await api.ocrEstimatePhotos(dataUrls)
    const lines = sheet.grid.map((row) => row[0] ?? '')
    const rates = parseCementSteelRateLines(lines)
    setManual((m) => ({
      ...m,
      cementRate: rates.cementRate ?? m.cementRate,
      steelRate: rates.steelRate ?? m.steelRate,
      msFlatsRate: rates.msFlatsRate ?? m.msFlatsRate,
      memoRef: rates.memoRef ?? m.memoRef,
      memoDate: rates.memoDate ?? m.memoDate,
      rateMonth: rates.monthYear ?? m.rateMonth
    }))
  }

  /** Cement & Steel rate circular upload (SE only) — OCR'd the same way as an
   * estimate photo (see src/pdfToImages.ts), then parsed for the TS Note's
   * rate sentence. Populates editable fields rather than baking values in
   * unreviewed, since OCR on a scanned circular can misread a digit. */
  async function uploadCementSteelRates(file: File) {
    setCementSteelBusy(true)
    setCementSteelError(null)
    try {
      const dataUrls = await pdfPagesToDataUrls(file)
      await fillCementSteelRatesFromPages(dataUrls)
      setCementSteelFileName(file.name)
    } catch (e) {
      setCementSteelError(e instanceof Error ? e.message : String(e))
    } finally {
      setCementSteelBusy(false)
    }
  }

  /** Auto-downloads the Cement & Steel rate circular for the month/year the
   * user picked, straight from the Public Health department site (the same
   * list the sidebar's Cement & Steel Rates tab browses), and fills the rate
   * fields from it — no manual upload needed when there's internet. On
   * failure (no internet, or no circular posted yet for that period), the
   * error points the user at the manual "Upload Cement & Steel rates" button
   * below, which works fully offline off an already-downloaded sheet. */
  async function fetchCementSteelRatesAuto() {
    const monthIndex = CEMENT_STEEL_MONTHS.indexOf(rateFetchMonth)
    if (monthIndex < 0 || !rateFetchYear.trim()) return
    setCementSteelBusy(true)
    setCementSteelError(null)
    try {
      const target = Number(rateFetchYear) * 12 + monthIndex
      const rates = await api.fetchCementSteelRates()
      const match = rates.find((r) => cementSteelRatePeriodKey(r) === target)
      if (!match) {
        throw new Error(
          `No Cement & Steel Rates circular found for ${rateFetchMonth} ${rateFetchYear} on the department site yet.`
        )
      }
      const bytes = await api.fetchCementSteelRateBytes(match.token)
      const dataUrls = await pdfPagesToDataUrlsFromData(bytes)
      await fillCementSteelRatesFromPages(dataUrls)
      setCementSteelFileName(match.description)
    } catch (e) {
      setCementSteelError(
        `${e instanceof Error ? e.message : String(e)} No internet, or use "Upload Cement & Steel rates" below to pick an already-downloaded sheet instead.`
      )
    } finally {
      setCementSteelBusy(false)
    }
  }

  // Optional non-responsiveness statement for the Bid Evaluation note — same
  // "List of Bidders Made Non-Responsive" sheet and reader the EE Note
  // Submitted flow uses (see WorkOrderAgreementTab's handleNonRespFile).
  async function handleNonRespFile(file: File) {
    setActionError(null)
    setBusy('nonresp')
    try {
      const lines = await pdfToPositionedLines(file)
      const { count } = summarizeNonResponsiveness(lines)
      setNonRespCount(count)
      setNonRespFileName(file.name)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function downloadIntimation(formats: ('docx' | 'pdf')[]) {
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillTemplate()
      const agencyName = notice?.agencyName ?? selectedRow?.['Name of the Agency']
      const tenderId = pdfEval?.tenderId?.replace(/[/\\]/g, '-')
      const label = seMode ? 'Letter of Acceptance' : 'Intimation'
      const name = [label, agencyName, tenderId].filter(Boolean).join(' ')
      const res = await api.exportCreatedDocument(filled, name, formats)
      setActionSaved(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
      if (res && res.length > 0) notifyIntimationIssued()
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function printIntimation() {
    setBusy('print')
    setActionError(null)
    try {
      const filled = await fillTemplate()
      const container = printScratchRef.current
      if (!container) throw new Error('Print failed to initialize.')
      await renderDocPreview(base64ToUint8(filled), container, { forSnapshot: true })
      await api.printCreatedDocument(container.innerHTML)
      notifyIntimationIssued()
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  const noWorks = !table || table.rows.length === 0

  // The SE fields no source can supply — grouped by which document(s) they
  // actually feed (matched against each template's own placeholder set), so
  // the expanded modal for a document shows only the fields it needs instead
  // of one big form for every document up front. Same `wo-date-row` /
  // `wo-date-field` markup as the Agreement page's own field rows, so fields
  // flow one after another (wrapping as needed) instead of a grid.
  const M = 'manual:'
  const mf = (key: keyof LoaManualFields, title: string, placeholder?: string): DocField => ({ label: M + key, title, placeholder })
  const itemNoField = mf('itemNo', 'Item No', 'e.g. 01')
  const periodField: DocField = { label: M + 'period', title: 'Period of Completion', periodUnit: true }
  const corpAbbr = (office.corporation ?? '').trim() || 'CMC'
  const asAuthorityFields: DocField[] = [
    {
      label: M + 'asAuthority',
      title: 'Admin Sanction Authority',
      options: [
        { value: 'commissioner', label: `Commissioner, ${corpAbbr}` },
        { value: 'zonal', label: 'Zonal Commissioner' },
        { value: 'resolution', label: 'Standing committee / resolution' }
      ]
    },
    ...(manual.asAuthority !== 'zonal' ? [mf('asDate', 'Admin Sanction Date', 'e.g. 21.07.2026')] : [])
  ]
  const tsNoDateFields: DocField[] = [mf('tsNo', 'TS No', 'e.g. 33'), mf('tsDate', 'TS Date', 'e.g. 04.07.2026')]
  const techBidOpenField = mf('techBidOpenDate', 'Technical Bid Opened Date', 'e.g. 06.08.2026')
  const bidEvalApprovedField = mf('bidEvalApprovedDate', 'Bid Evaluation Approved Date', 'e.g. 06.08.2026')

  // The SE fields no source can supply, grouped per document (matched to each
  // template's own placeholder set) and shown in the same fields-left /
  // live-page-right editor the Intimation letter uses. Labels carry the
  // `manual:` prefix so edits route to `manual` rather than the letter's overrides.
  const seManualGroups: Partial<Record<ExpandedDoc, DocFieldGroup[]>> = {
    tsNote: [
      { title: 'References', fields: [...tsNoDateFields, mf('ssrYear', 'SSR Year', 'e.g. 2026-27'), ...asAuthorityFields] },
      {
        title: 'Rates',
        fields: [
          mf('cementRate', 'Cement Rate (Rs./MT)', 'e.g. 5,100'),
          mf('steelRate', 'Steel TMT/HYSD Rate (Rs./MT)', 'e.g. 50,000'),
          mf('msFlatsRate', 'M.S. Flats Rate (Rs./MT)', 'e.g. 55,000'),
          mf('memoRef', 'Memo Ref', 'e.g. 146/Cement & Steel/T3/May/2025-26'),
          mf('memoDate', 'Memo Date', 'e.g. 29.05.2026'),
          mf('rateMonth', 'Rate Month', 'e.g. May-2026')
        ]
      }
    ],
    eligibility: [
      { title: 'The work', fields: [itemNoField, periodField, ...asAuthorityFields, ...tsNoDateFields] },
      {
        title: 'Eligibility clauses',
        fields: [
          mf('classOfRegistration', 'Class of Registration'),
          mf('emdValidity', 'EMD Validity', 'e.g. 6 Months'),
          mf('solvencyValidity', 'Solvency Certificate Validity', 'e.g. 1 year'),
          mf('graduateEngineers', 'Graduate Engineers (Civil)', 'e.g. 1 No.'),
          mf('diplomaEngineers', 'Diploma Engineers (Civil)', 'e.g. 2 Nos.'),
          mf('equipment1', 'Critical Equipment 1'),
          mf('equipment2', 'Critical Equipment 2'),
          mf('equipment3', 'Critical Equipment 3'),
          mf('goMsNo', 'G.O. MS No.', 'e.g. 66')
        ]
      }
    ],
    bidEval: [{ title: 'References and dates', fields: [...asAuthorityFields, ...tsNoDateFields, itemNoField, techBidOpenField] }],
    agencyApproval: [
      { title: 'References and dates', fields: [...asAuthorityFields, ...tsNoDateFields, itemNoField, techBidOpenField, bidEvalApprovedField] }
    ]
  }
  // The SE Letter of Acceptance's fields in the order the letter reads, with plain
  // names — instead of the generic bucketing the EE letter gets. Item No and Period
  // are the hand-typed `manual` fields (so they also feed the companion notes), not
  // the template's own copies of those placeholders, which would only edit this letter.
  const seIntimationGroups: DocFieldGroup[] = (() => {
    const shown = new Set<string>()
    const pick = (label: string, title: string, extra: Partial<DocField> = {}): DocField[] => {
      if (!labels.includes(label)) return []
      shown.add(label)
      return [{ label, title, ...extra }]
    }
    const hidden = ['Item No', 'Period of Completion', 'Period Completion']
    hidden.forEach((l) => shown.add(l))
    const groups: DocFieldGroup[] = [
      {
        title: 'Intimation',
        fields: [
          ...pick('Intimation Lr No', 'Intimation Lr. No', { wide: true }),
          ...pick('LOA Date', 'Intimation Date'),
          itemNoField,
          periodField
        ]
      },
      {
        title: 'Agency',
        fields: [
          ...pick('Agency Name', 'Agency name', { wide: true }),
          ...pick('Address of the agency', 'Address', { multiline: true }),
          ...pick('agency phone number', 'Phone number')
        ]
      },
      {
        title: 'The work',
        fields: [
          ...pick('Name of the work', 'Name of the work', { multiline: true }),
          ...pick('Reserved Tag', 'Reservation tag', { placeholder: 'e.g. (Reserved for SC category)' }),
          ...pick('Circle Line', 'Circle'),
          ...pick('Nit No', 'NIT No'),
          ...pick('Price Bid opening date', 'Price bid opened on'),
          mf('adminSanction', 'Admin. Sanction Value (Rs.)', 'e.g. 16850000')
        ]
      },
      {
        title: 'Amounts',
        fields: [
          ...pick('ECV', 'Estimate Contract Value (Rs.)'),
          ...pick('Tender Percentage', 'Tender percentage (% below)'),
          ...pick('Contract Amount', 'Contract amount (Rs.)'),
          ...pick('Contract In Words', 'Contract amount in words', { multiline: true }),
          ...pick('MMC Emd Clause', 'EMD clause', { multiline: true }),
          ...pick('MMC Reserved Emd Clause', 'EMD clause', { multiline: true }),
          ...pick('BG Validity Months', 'BG validity (months)'),
          ...pick('Period Number', 'Period (number only)')
        ]
      },
      {
        title: 'Office',
        fields: [...pick('Zone', 'Zone'), ...pick('Corporation', 'Corporation')]
      }
    ]
    // Anything the template carries that isn't listed above still gets a field.
    const rest = labels.filter((l) => !shown.has(l))
    if (rest.length > 0) groups.push({ title: 'Other', fields: rest.map((l) => ({ label: l, title: humanizeLabel(l) })) })
    return groups.filter((g) => g.fields.length > 0)
  })()
  const manualDerived: Record<string, string> = Object.fromEntries(
    Object.entries(manual).map(([k, v]) => [M + k, String(v ?? '')])
  )

  function setManualField(label: string, value: string) {
    setManual((m) => ({ ...m, [label.slice(M.length)]: value }) as LoaManualFields)
  }

  const bundleProgressLabel = bundleProgress
    ? `${bundleProgress.phase === 'preparing' ? 'Preparing' : 'Saving'} ${bundleProgress.done} of ${bundleProgress.total}…`
    : 'Preparing…'

  const downloadAllButton = seMode && (
    <div className="download-all-group">
      <div className="format-toggle" role="group" aria-label="Download format">
        <button
          type="button"
          className={bundleFormat === 'docx' ? 'on' : ''}
          disabled={busy === 'bundle'}
          onClick={() => setBundleFormat('docx')}
        >
          Word
        </button>
        <button
          type="button"
          className={bundleFormat === 'pdf' ? 'on' : ''}
          disabled={busy === 'bundle'}
          onClick={() => setBundleFormat('pdf')}
        >
          PDF
        </button>
      </div>
      <button
        className="primary"
        disabled={busy === 'bundle'}
        onClick={downloadAll}
        title={`Letter of Acceptance, TS Note, Eligibility Criteria, Bid Evaluation & Agency Approval — all as ${bundleFormat === 'pdf' ? 'PDF' : 'Word'}, into one folder you choose`}
      >
        <IconDownload /> {busy === 'bundle' ? bundleProgressLabel : 'Download all documents'}
      </button>
    </div>
  )

  // Verify only once the source documents that feed it are actually uploaded
  // and the letter has something generated to check — same "required
  // documents uploaded and document is generated" gate as the rest of the tab.
  const verifyReady = !!notice && !!pdfEval && showBody && !!templateB64
  const headerActions = (
    <>
      {downloadAllButton}
      {verifyReady && <VerifyButton getItems={getVerifyItems} />}
    </>
  )

  return (
    <div className="card">
      {headerMounted && headerActionRef?.current && createPortal(headerActions, headerActionRef.current)}
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <div className="empty empty--tight">
        <IconBell />
        {seMode && (
          <div className="gps-toolbar cement-steel-fetch">
            <label className="wo-date-field">
              <span>Steel Rate Month</span>
              <select value={rateFetchMonth} onChange={(e) => setRateFetchMonth(e.target.value)} disabled={cementSteelBusy}>
                <option value="">Month</option>
                {CEMENT_STEEL_MONTHS.map((m) => (
                  <option key={m} value={m}>
                    {m.charAt(0).toUpperCase() + m.slice(1)}
                  </option>
                ))}
              </select>
            </label>
            <label className="wo-date-field">
              <span>Year</span>
              <input
                type="text"
                inputMode="numeric"
                placeholder="e.g. 2026"
                value={rateFetchYear}
                onChange={(e) => setRateFetchYear(e.target.value)}
                disabled={cementSteelBusy}
              />
            </label>
            <button
              type="button"
              className="primary upload-btn"
              onClick={() => void fetchCementSteelRatesAuto()}
              disabled={cementSteelBusy || !rateFetchMonth || !rateFetchYear.trim()}
              title="Downloads the circular for this month/year from the Public Health department site — same list as the sidebar's Cement & Steel Rates tab"
            >
              <IconDownload /> {cementSteelBusy ? 'Fetching…' : 'Fetch Cement & Steel rates'}
            </button>
          </div>
        )}
        <div className="boq-actions boq-actions--grid">
          <button
            className="primary upload-btn"
            onClick={() => noticeInputRef.current?.click()}
            disabled={!templateB64 || busy === 'notice'}
          >
            <IconFolder /> {busy === 'notice' ? 'Reading…' : notice ? 'Change Online Intimation' : 'Upload Online Intimation'}
          </button>
          <button
            className="primary upload-btn"
            onClick={() => supplierDetailsInputRef.current?.click()}
            disabled={!templateB64 || busy === 'supplier'}
            title="eGP portal's 'Supplier Details' page — fills the agency's phone number, which the Online Intimation letter never carries"
          >
            <IconFolder />{' '}
            {busy === 'supplier' ? 'Reading…' : supplierDetailsName ? 'Change Supplier Details' : 'Upload Supplier Details'}
          </button>
          <button
            className="primary upload-btn"
            onClick={() => pdfInputRef.current?.click()}
            disabled={!templateB64 || busy === 'pdf'}
          >
            <IconFolder /> {busy === 'pdf' ? 'Reading PDF…' : pdfEval ? 'Change L1 selection form' : 'Upload L1 selection form'}
          </button>
          {seMode && (
            <button
              className="primary upload-btn"
              onClick={() => cementSteelInputRef.current?.click()}
              disabled={cementSteelBusy}
            >
              <IconFolder />{' '}
              {cementSteelBusy ? 'Reading…' : cementSteelFileName ? 'Change Cement & Steel rates' : 'Upload Cement & Steel rates'}
            </button>
          )}
          {seMode && (
            <button
              className="primary upload-btn"
              onClick={() => nonRespInputRef.current?.click()}
              disabled={!pdfEval || busy === 'nonresp'}
            >
              <IconFolder /> {busy === 'nonresp' ? 'Reading…' : nonRespFileName ? 'Change Non-responsive form' : 'Upload Non-responsive form'}
            </button>
          )}
          <input
            ref={noticeInputRef}
            type="file"
            accept=".html,.htm,text/html,.pdf,application/pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handleNoticeFile(file)
              e.target.value = ''
            }}
          />
          <input
            ref={supplierDetailsInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handleSupplierDetailsFile(file)
              e.target.value = ''
            }}
          />
          <input
            ref={pdfInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handlePdfFile(file)
              e.target.value = ''
            }}
          />
          <input
            ref={cementSteelInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void uploadCementSteelRates(file)
              e.target.value = ''
            }}
          />
          <input
            ref={nonRespInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handleNonRespFile(file)
              e.target.value = ''
            }}
          />
        </div>
        {noticeName && <p className="estimate-hint">Address read from {noticeName}</p>}
        {supplierDetailsName && <p className="estimate-hint">Phone number/address read from {supplierDetailsName}</p>}
        {pdfName && <p className="estimate-hint">Tender details read from {pdfName}</p>}
        {cementSteelFileName && (
          <p className="estimate-hint">
            Cement/Steel/M.S.Flats rates read from {cementSteelFileName} — used in the TS Note; check them below before generating.
          </p>
        )}
        {nonRespFileName && (
          <p className="estimate-hint">
            Non-responsive bidders read from {nonRespFileName}: ({nonRespCount}) of ({allBidders.length}) — used in the Bid Evaluation note.
          </p>
        )}
        {pdfStatus && (
          <div className="notice ok">
            <IconBell /> {pdfStatus}
          </div>
        )}
        {cementSteelError && (
          <div className="notice error">
            <IconWarn /> {cementSteelError}
          </div>
        )}
      </div>

      {loadError && (
        <div className="notice error">
          <IconWarn /> {loadError}
        </div>
      )}
      {workMismatch && workMatch && (
        <div className="notice error">
          <IconWarn /> {sameWorkMismatchMessage(workMatch)}
        </div>
      )}
      {noWorksRowMatch && pdfEval && (
        <div className="notice">
          <IconWarn /> “{pdfEval.nameOfWork}” isn’t in your Works List, so its supporting details (Circle, etc.) are left
          blank — the letter still fills the name of work and amounts from the uploaded L1 / Intimation. Add it to the
          Works List if you want those details filled in automatically.
        </div>
      )}

      {showBody ? (
        <>
          {noWorks && (
            <div className="notice">
              Works List is empty — supporting details (Circle, etc.) will be blank until you add the work there; the
              documents below still fill their name of work and amounts from your uploads.
            </div>
          )}
          {actionSaved && !expanded && (
            <div className="notice ok">
              <IconCheck /> {actionSaved}
            </div>
          )}
          {actionError && !expanded && (
            <div className="notice error">
              <IconWarn /> {actionError}
            </div>
          )}

          {!headerMounted && downloadAllButton}

          <div className={`wo-tiles${seMode && seDocsReady ? '' : ' wo-tiles-single'}`}>
            <button className="wo-tile" onClick={() => openDoc('intimation')}>
              <div className="wo-tile-preview">
                <div ref={previewRef} className="wo-tile-doc" />
                <span className="wo-tile-open">Click to preview</span>
              </div>
              <div className="wo-tile-foot">{DOC_LABEL.intimation}</div>
            </button>

            {seMode && seDocsReady && (
              <>
                <button className="wo-tile" onClick={() => openDoc('tsNote')}>
                  <div className="wo-tile-preview">
                    <div ref={tsNotePreviewRef} className="wo-tile-doc" />
                    <span className="wo-tile-open">Click to preview</span>
                  </div>
                  <div className="wo-tile-foot">{DOC_LABEL.tsNote}</div>
                </button>
                <button className="wo-tile" onClick={() => openDoc('eligibility')}>
                  <div className="wo-tile-preview">
                    <div ref={eligibilityPreviewRef} className="wo-tile-doc" />
                    <span className="wo-tile-open">Click to preview</span>
                  </div>
                  <div className="wo-tile-foot">{DOC_LABEL.eligibility}</div>
                </button>
                <button className="wo-tile" onClick={() => openDoc('bidEval')}>
                  <div className="wo-tile-preview">
                    <div ref={bidEvalPreviewRef} className="wo-tile-doc" />
                    <span className="wo-tile-open">Click to preview</span>
                  </div>
                  <div className="wo-tile-foot">{DOC_LABEL.bidEval}</div>
                </button>
                <button className="wo-tile" onClick={() => openDoc('agencyApproval')}>
                  <div className="wo-tile-preview">
                    <div ref={agencyApprovalPreviewRef} className="wo-tile-doc" />
                    <span className="wo-tile-open">Click to preview</span>
                  </div>
                  <div className="wo-tile-foot">{DOC_LABEL.agencyApproval}</div>
                </button>
              </>
            )}
          </div>
        </>
      ) : noWorks ? (
        <div className="notice">
          Add works to the Works List first, or upload the Online Intimation and L1 selection form above — either fills
          the Intimation letter.
        </div>
      ) : null}

      {expanded &&
        createPortal(
          <div className="wo-modal-overlay" onClick={() => setExpanded(null)}>
          <div
            className="wo-modal wide doc-edit-modal"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="wo-modal-head">
              <span className="wo-modal-title">
                {DOC_LABEL[expanded]}
                {expandedPages > 1 ? ` — ${expandedPages} pages` : ''}
              </span>
              <button className="wo-modal-close" onClick={() => setExpanded(null)} title="Close" aria-label="Close">
                ×
              </button>
            </div>

            <div className="wo-modal-main">
              <div className="wo-modal-body">
                {/* Every document: its fields on the left, the live page on the
                    right — same editor split the Agreement page's documents use. */}
                <div className="ns-body">
                  <DocFieldsEditor
                    groups={seMode && expanded === 'intimation' ? seIntimationGroups : [...(expanded === 'intimation' ? intimationGroups : []), ...(seMode ? seManualGroups[expanded] ?? [] : [])]}
                    derived={{ ...values, ...manualDerived }}
                    edits={intimationEdits}
                    onChange={(label, value) =>
                      label.startsWith(M) ? setManualField(label, value) : setIntimationEdits((prev) => ({ ...prev, [label]: value }))
                    }
                    onRevert={(label) =>
                      setIntimationEdits((prev) => {
                        const next = { ...prev }
                        delete next[label]
                        return next
                      })
                    }
                    onResetAll={() => setIntimationEdits({})}
                  />
                  <div className="ns-preview">
                    <div className="ns-preview-scroll">
                      <div className="page-margin-host">
                        <div ref={expandedRef} className="intimation-docx-preview" />
                        <PageMarginGuides
                          containerRef={expandedRef}
                          margins={docMargins[expanded] ?? null}
                          onChange={(m) => setMarginsFor(expanded, m)}
                          renderKey={expandedPages}
                        />
                        <button
                          type="button"
                          className="ghost page-margin-reset"
                          disabled={!docMargins[expanded]}
                          onClick={() => setMarginsFor(expanded, null)}
                        >
                          Reset margins
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="wo-modal-foot">
              {actionError && (
                <div className="notice error" style={{ marginRight: 'auto' }}>
                  <IconWarn /> {actionError}
                </div>
              )}
              <button className="ghost wo-foot-close" onClick={() => setExpanded(null)}>
                Close
              </button>
              {expanded === 'intimation' ? (
                <>
                  <button className="primary" onClick={() => downloadIntimation(['docx'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
                  </button>
                  <button className="primary" onClick={() => downloadIntimation(['pdf'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'pdf' ? 'Saving…' : 'PDF'}
                  </button>
                  <button className="primary" onClick={printIntimation} disabled={busy !== null}>
                    <IconPrint /> {busy === 'print' ? 'Opening…' : 'Print'}
                  </button>
                </>
              ) : (
                <>
                  <button className="primary" onClick={() => downloadSeDoc(expanded, ['docx'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
                  </button>
                  <button className="primary" onClick={() => downloadSeDoc(expanded, ['pdf'])} disabled={busy !== null}>
                    <IconDownload /> {busy === 'pdf' ? 'Saving…' : 'PDF'}
                  </button>
                  <button className="primary" onClick={() => printSeDoc(expanded)} disabled={busy !== null}>
                    <IconPrint /> {busy === 'print' ? 'Opening…' : 'Print'}
                  </button>
                </>
              )}
            </div>
          </div>
          </div>,
          document.body
        )}
    </div>
  )
}

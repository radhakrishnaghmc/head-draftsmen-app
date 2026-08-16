import { useEffect, useLayoutEffect, useMemo, useRef, useState, type ReactNode, type RefObject } from 'react'
import { createPortal } from 'react-dom'
import { renderAsync } from '../lazyDocxPreview'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, parseAllBidders, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import { computeWorkAmounts, tenderPercentFromRow } from '@core/worksAmounts'
import { wrapAgencyAddress } from '@core/workOrderAgreement'
import { stripItemNoTag } from '@core/bidDocument'
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
import { resolveFromDirectory, entriesOf } from '../zoneCircleDirectory'
import type { Office } from '../office'
import type { PlaceholderMatch } from '@core/createDocument'
import { pdfToTextLines, pdfToPositionedLines } from '../pdfToText'
import { pdfPagesToDataUrls } from '../pdfToImages'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH, normalizeDocxTextboxes } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell, IconCheck } from './Icons'
import type { ExcelTable } from '@core/types'
import type { AgreementBundleFile } from '../../electron/ipc-contract'

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
  asAuthority: 'zonal' | 'commissioner'
  /** Only used when asAuthority is 'commissioner' — the zonal form's Ref line carries no date in the office's own samples. */
  asDate: string
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
}

const LOA_MANUAL_DEFAULTS: LoaManualFields = {
  adminSanction: '',
  period: '3 Months',
  itemNo: '',
  loaDate: '',
  asAuthority: 'commissioner',
  asDate: '',
  tsNo: '',
  tsDate: '',
  techBidOpenDate: '',
  bidEvalApprovedDate: '',
  cementRate: '',
  steelRate: '',
  msFlatsRate: '',
  memoRef: '',
  memoDate: '',
  rateMonth: ''
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
function asAuthorityText(zoneAbbrCode: string, m: LoaManualFields): string {
  return m.asAuthority === 'zonal' ? `Zonal Commissioner, ${zoneAbbrCode}, CMC` : 'Commissioner, CMC'
}

/** The Ref-1 sentence tail: the authority, plus ", dated:DD.MM.YYYY" only for the Commissioner form — the zonal form's Ref line carries no date in the office's own samples. */
function asRefLineText(zoneAbbrCode: string, m: LoaManualFields): string {
  const authority = asAuthorityText(zoneAbbrCode, m)
  return m.asAuthority === 'zonal' ? authority : `${authority}, dated:${m.asDate.trim() || '__________'}`
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

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
      return row['Phone number of the agency'] ?? ''
    case 'item no':
      return manual.itemNo
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
      return ecv != null ? formatIndianAmount(Math.round(ecv * 0.0004), 0) : ''
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
  const estimateLakhs = (row['Amount of estimate'] ?? '').replace(/,/g, '').trim()

  const key = norm(label)
  switch (key) {
    case 'zone':
      return office.zone ?? ''
    case 'zone abbr':
      return zoneAbbrCode
    case 'financial year':
      return financialYearFromDate(pdf.noticeDate)
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
      return asAuthorityText(zoneAbbrCode, manual)
    case 'as ref line':
      return asRefLineText(zoneAbbrCode, manual)
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
  return {
    zoneAbbr: zoneAbbrCode,
    workName: work.name,
    estimateLakhs: (row['Amount of estimate'] ?? '').replace(/,/g, '').trim(),
    asAuthority: asAuthorityText(zoneAbbrCode, manual),
    asRefLine: asRefLineText(zoneAbbrCode, manual),
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
export default function GiveIntimationTab({ tables, onChange, office, headerActionRef }: Props) {
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

  const [templateB64, setTemplateB64] = useState<string | null>(null)
  const [labels, setLabels] = useState<string[]>([])
  const [loadError, setLoadError] = useState<string | null>(null)
  const [manual, setManual] = useState<LoaManualFields>(LOA_MANUAL_DEFAULTS)

  const [rowIndex, setRowIndex] = useState(0)
  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  const [pdfEval, setPdfEval] = useState<TenderEvaluation | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pdfStatus, setPdfStatus] = useState<string | null>(null)
  // Whether the uploaded L1 form's work matched a Works List row. false means
  // it matched none, so the letter would fill from an unrelated row — the
  // preview is gated until the work is added to the Works List. null = not yet
  // determined (no L1 uploaded, or no works to match against).
  const [worksRowMatched, setWorksRowMatched] = useState<boolean | null>(null)
  const [values, setValues] = useState<Record<string, string>>({})

  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf' | 'bundle'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  // Which format "Download all documents" saves every document as. Defaults
  // to Word since PDF needs LibreOffice installed.
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
  // From the uploaded "List of Bidders Made Non-Responsive" sheet — feeds the
  // Bid Evaluation note's participated/responsive/non-responsive counts, the
  // same way the EE Note Submitted's non-responsive upload works.
  const [nonRespCount, setNonRespCount] = useState(0)
  const [nonRespFileName, setNonRespFileName] = useState('')

  const noticeInputRef = useRef<HTMLInputElement>(null)
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
  const detailsRow = !notice || !pdfEval ? {} : worksRowMatched === false ? {} : (selectedRow ?? {})

  // A reserved (SC/ST) work uses the LOA variant that omits the EMD balance item.
  const seReserved = seMode && isReservedWork(detailsRow, pdfEval ?? {})

  // Load the bundled format for this office/work (SE LOA — reserved or not —
  // when zone-level, otherwise the EE Intimation) and read its placeholders.
  // Reloads if the office kind flips or the selected work's reserved status does.
  useEffect(() => {
    let cancelled = false
    setTemplateB64(null)
    setLabels([])
    void (async () => {
      try {
        const b64 = seMode ? await api.loaSeTemplate(seReserved) : await api.intimationTemplate()
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
  }, [seMode, seReserved])

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
          ? resolveLoaValue(label, notice ?? {}, pdfEval ?? {}, detailsRow, office, manual)
          : resolveIntimationValue(label, notice ?? {}, pdfEval ?? {}, detailsRow, office)
      return next
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [labels, rowIndex, notice, pdfEval, table, seMode, office, manual, worksRowMatched])

  // The Online Intimation can be uploaded as either the portal "View Intimation
  // Notice" .html page or the printed Intimation / LOA .pdf — both carry the
  // agency, address, NIT No, ECV and contract value, so parse by file type.
  async function handleNoticeFile(file: File) {
    setActionError(null)
    try {
      const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
      const parsed = isPdf
        ? parseIntimationNoticeText(await pdfToTextLines(file))
        : parseIntimationNotice(await file.text())
      setNotice(parsed)
      setNoticeName(file.name)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    }
  }

  // The L-1 selection / evaluation PDF fills the letter's tender fields and
  // updates that work's Works List row (matched by Name of Work), auto-selecting it.
  async function handlePdfFile(file: File) {
    setBusy('pdf')
    setActionError(null)
    setPdfStatus(null)
    try {
      const lines = await pdfToTextLines(file)
      const ev = parseTenderEvaluation(lines)
      if (!ev.nameOfWork && !ev.tenderId) {
        throw new Error("Couldn't read tender details from that PDF — is it the Commercial Evaluation / Stage Selected page?")
      }
      setPdfEval(ev)
      setPdfName(file.name)
      setWorksRowMatched(null)
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
        // Match only — the Works List database is updated solely from the Works
        // List page ("Update from L1"), never here. We just find the row so its
        // supporting details fill the letter and select it (an embedding /
        // wording-drift match has no exact name to re-derive from, so without
        // this the selection would stay on row 0 and fill a different work).
        const { matchedCount, matchedRowIndices } = updateWorksListFromEvaluations(table, [ev], embeddings)
        if (matchedCount > 0) {
          const idx = matchedRowIndices[0]
          if (idx != null && idx >= 0) setRowIndex(idx)
          setWorksRowMatched(true)
          setPdfStatus(`Matched "${ev.nameOfWork}" to a Works List row — its details fill the letter.`)
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
    return api.fillPlaceholdersInDocument(templateB64, resolved, values)
  }

  // Guard: the Online Intimation and the L1 selection form must describe the
  // same work (matched by NIT No, or agency name when the NIT No is absent).
  const workMatch = useMemo(() => (notice && pdfEval ? checkSameWork(notice, pdfEval) : null), [notice, pdfEval])
  const workMismatch = workMatch?.status === 'mismatch'
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

  // Live docx thumbnail of the filled letter (tile preview) — refreshed whenever the values change.
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
  }, [templateB64, values, showBody])

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
      ...seRefsFor(pdfEval ?? {}, notice ?? {}, detailsRow, office, manual),
      bidders: allBidders,
      nonRespCount
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [pdfEval, notice, detailsRow, office, manual, allBidders, nonRespCount]
  )
  const agencyApprovalData: AgencyApprovalData = useMemo(
    () => ({
      ...seRefsFor(pdfEval ?? {}, notice ?? {}, detailsRow, office, manual),
      zone: office.zone ?? '',
      // The L-1 sheet's own Server Time is also when the Technical Bid
      // evaluation was approved (same sitting), same fallback as techBidOpenDate.
      bidEvalApprovedDate: pdfEval?.serverDate || manual.bidEvalApprovedDate.trim(),
      bidders: allBidders
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [pdfEval, notice, detailsRow, office, manual, allBidders]
  )

  async function fillSeDoc(kind: 'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval'): Promise<string> {
    if (kind === 'bidEval') return api.noteSubmittedDocx(buildBidEvaluationHtml(bidEvalData))
    if (kind === 'agencyApproval') return api.noteSubmittedDocx(buildAgencyApprovalHtml(agencyApprovalData))
    const b64 = seDocB64[kind]
    if (!b64) throw new Error('Format not loaded yet.')
    const kindLabels = seDocLabels[kind] ?? []
    const resolved: PlaceholderMatch[] = kindLabels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(b64, resolved, seDocValues[kind] ?? {})
  }

  async function renderSeDocInto(kind: 'tsNote' | 'eligibility' | 'bidEval' | 'agencyApproval', container: HTMLElement) {
    const filled = await fillSeDoc(kind)
    container.innerHTML = ''
    await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
    normalizeDocxTextboxes(container)
  }

  // Live previews for the 4 SE companion notes — refreshed whenever their inputs change.
  useEffect(() => {
    if (!seDocsReady) return
    if (tsNotePreviewRef.current) void renderSeDocInto('tsNote', tsNotePreviewRef.current).catch(() => {})
    if (eligibilityPreviewRef.current) void renderSeDocInto('eligibility', eligibilityPreviewRef.current).catch(() => {})
    if (bidEvalPreviewRef.current) void renderSeDocInto('bidEval', bidEvalPreviewRef.current).catch(() => {})
    if (agencyApprovalPreviewRef.current) void renderSeDocInto('agencyApproval', agencyApprovalPreviewRef.current).catch(() => {})
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
    void (async () => {
      try {
        if (expanded === 'intimation') {
          const filled = await fillTemplate()
          container.innerHTML = ''
          await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
          normalizeDocxTextboxes(container)
        } else {
          await renderSeDocInto(expanded, container)
        }
        setExpandedPages(container.querySelectorAll('section.docx').length)
      } catch (e) {
        setActionError(e instanceof Error ? e.message : String(e))
      }
    })()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [expanded, values, seDocValues, bidEvalData, agencyApprovalData])

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
      container.innerHTML = ''
      await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
      normalizeDocxTextboxes(container)
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
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
            `PDFs need LibreOffice installed — install it from libreoffice.org, or download those as Word (.docx) instead.`
        )
      } else {
        setActionSaved(`Saved ${res.written.length} document(s) to the chosen folder.`)
      }
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      unsubscribe()
      setBundleProgress(null)
      setBusy(null)
    }
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
      const sheet = await api.ocrEstimatePhotos(dataUrls)
      const lines = sheet.grid.map((row) => row[0] ?? '')
      const rates = parseCementSteelRateLines(lines)
      setCementSteelFileName(file.name)
      setManual((m) => ({
        ...m,
        cementRate: rates.cementRate ?? m.cementRate,
        steelRate: rates.steelRate ?? m.steelRate,
        msFlatsRate: rates.msFlatsRate ?? m.msFlatsRate,
        memoRef: rates.memoRef ?? m.memoRef,
        memoDate: rates.memoDate ?? m.memoDate,
        rateMonth: rates.monthYear ?? m.rateMonth
      }))
    } catch (e) {
      setCementSteelError(e instanceof Error ? e.message : String(e))
    } finally {
      setCementSteelBusy(false)
    }
  }

  // Optional non-responsiveness statement for the Bid Evaluation note — same
  // "List of Bidders Made Non-Responsive" sheet and reader the EE Note
  // Submitted flow uses (see WorkOrderAgreementTab's handleNonRespFile).
  async function handleNonRespFile(file: File) {
    setActionError(null)
    try {
      const lines = await pdfToPositionedLines(file)
      const { count } = summarizeNonResponsiveness(lines)
      setNonRespCount(count)
      setNonRespFileName(file.name)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    }
  }

  async function downloadIntimation(formats: ('docx' | 'pdf')[]) {
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillTemplate()
      const agencyName = notice?.agencyName ?? selectedRow?.['Name of the Agency']
      const name = `${seMode ? 'Letter of Acceptance' : 'Intimation'}${agencyName ? ` - ${agencyName}` : ''}`
      const res = await api.exportCreatedDocument(filled, name, formats)
      setActionSaved(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
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
      container.innerHTML = ''
      await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
      normalizeDocxTextboxes(container)
      await api.printCreatedDocument(container.innerHTML)
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
  const itemNoField = (
    <label className="wo-date-field">
      <span>Item No</span>
      <input type="text" placeholder="e.g. 01" value={manual.itemNo} onChange={(e) => setManual((m) => ({ ...m, itemNo: e.target.value }))} />
    </label>
  )
  const periodField = (
    <label className="wo-date-field">
      <span>Period of Completion</span>
      <input type="text" placeholder="e.g. 3 Months" value={manual.period} onChange={(e) => setManual((m) => ({ ...m, period: e.target.value }))} />
    </label>
  )
  const asAuthorityFields = (
    <>
      <label className="wo-date-field">
        <span>Admin Sanction Authority</span>
        <select value={manual.asAuthority} onChange={(e) => setManual((m) => ({ ...m, asAuthority: e.target.value as 'zonal' | 'commissioner' }))}>
          <option value="commissioner">Commissioner, CMC</option>
          <option value="zonal">Zonal Commissioner</option>
        </select>
      </label>
      {manual.asAuthority === 'commissioner' && (
        <label className="wo-date-field">
          <span>Admin Sanction Date</span>
          <input type="text" placeholder="e.g. 21.07.2026" value={manual.asDate} onChange={(e) => setManual((m) => ({ ...m, asDate: e.target.value }))} />
        </label>
      )}
    </>
  )
  const tsNoDateFields = (
    <>
      <label className="wo-date-field">
        <span>TS No</span>
        <input type="text" placeholder="e.g. 33" value={manual.tsNo} onChange={(e) => setManual((m) => ({ ...m, tsNo: e.target.value }))} />
      </label>
      <label className="wo-date-field">
        <span>TS Date</span>
        <input type="text" placeholder="e.g. 04.07.2026" value={manual.tsDate} onChange={(e) => setManual((m) => ({ ...m, tsDate: e.target.value }))} />
      </label>
    </>
  )
  const techBidOpenField = (
    <label className="wo-date-field">
      <span>Technical Bid Opened Date</span>
      <input
        type="text"
        placeholder="e.g. 06.08.2026"
        value={manual.techBidOpenDate}
        onChange={(e) => setManual((m) => ({ ...m, techBidOpenDate: e.target.value }))}
      />
    </label>
  )
  const bidEvalApprovedField = (
    <label className="wo-date-field">
      <span>Bid Evaluation Approved Date</span>
      <input
        type="text"
        placeholder="e.g. 06.08.2026"
        value={manual.bidEvalApprovedDate}
        onChange={(e) => setManual((m) => ({ ...m, bidEvalApprovedDate: e.target.value }))}
      />
    </label>
  )

  // Fields grouped per document, matched to each template's actual placeholder set.
  const modalFields: Partial<Record<ExpandedDoc, ReactNode>> = {
    intimation: (
      <>
        <label className="wo-date-field">
          <span>Admin. Sanction Value</span>
          <input
            type="text"
            placeholder="e.g. 16850000"
            value={manual.adminSanction}
            onChange={(e) => setManual((m) => ({ ...m, adminSanction: e.target.value }))}
          />
        </label>
        {periodField}
        {itemNoField}
        <label className="wo-date-field">
          <span>LOA Date</span>
          <input type="text" placeholder="e.g. 20.03.2026" value={manual.loaDate} onChange={(e) => setManual((m) => ({ ...m, loaDate: e.target.value }))} />
        </label>
      </>
    ),
    tsNote: (
      <>
        {tsNoDateFields}
        {asAuthorityFields}
        <label className="wo-date-field">
          <span>Cement Rate (Rs./MT)</span>
          <input type="text" placeholder="e.g. 5,100" value={manual.cementRate} onChange={(e) => setManual((m) => ({ ...m, cementRate: e.target.value }))} />
        </label>
        <label className="wo-date-field">
          <span>Steel TMT/HYSD Rate (Rs./MT)</span>
          <input type="text" placeholder="e.g. 50,000" value={manual.steelRate} onChange={(e) => setManual((m) => ({ ...m, steelRate: e.target.value }))} />
        </label>
        <label className="wo-date-field">
          <span>M.S. Flats Rate (Rs./MT)</span>
          <input type="text" placeholder="e.g. 55,000" value={manual.msFlatsRate} onChange={(e) => setManual((m) => ({ ...m, msFlatsRate: e.target.value }))} />
        </label>
        <label className="wo-date-field">
          <span>Memo Ref</span>
          <input
            type="text"
            placeholder="e.g. 146/Cement & Steel/T3/May/2025-26"
            value={manual.memoRef}
            onChange={(e) => setManual((m) => ({ ...m, memoRef: e.target.value }))}
          />
        </label>
        <label className="wo-date-field">
          <span>Memo Date</span>
          <input type="text" placeholder="e.g. 29.05.2026" value={manual.memoDate} onChange={(e) => setManual((m) => ({ ...m, memoDate: e.target.value }))} />
        </label>
        <label className="wo-date-field">
          <span>Rate Month</span>
          <input type="text" placeholder="e.g. May-2026" value={manual.rateMonth} onChange={(e) => setManual((m) => ({ ...m, rateMonth: e.target.value }))} />
        </label>
      </>
    ),
    eligibility: (
      <>
        {itemNoField}
        {periodField}
        {asAuthorityFields}
        {tsNoDateFields}
      </>
    ),
    bidEval: (
      <>
        {asAuthorityFields}
        {tsNoDateFields}
        {itemNoField}
        {techBidOpenField}
      </>
    ),
    agencyApproval: (
      <>
        {asAuthorityFields}
        {tsNoDateFields}
        {itemNoField}
        {techBidOpenField}
        {bidEvalApprovedField}
      </>
    )
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

  return (
    <div className="card">
      {headerMounted && headerActionRef?.current && createPortal(downloadAllButton, headerActionRef.current)}
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <div className="empty empty--tight">
        <IconBell />
        <div className="boq-actions boq-actions--grid">
          <button className="primary upload-btn" onClick={() => noticeInputRef.current?.click()} disabled={!templateB64}>
            <IconFolder /> {notice ? 'Change Online Intimation' : 'Upload Online Intimation'}
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
            <button className="primary upload-btn" onClick={() => nonRespInputRef.current?.click()} disabled={!pdfEval}>
              <IconFolder /> {nonRespFileName ? 'Change Non-responsive form' : 'Upload Non-responsive form'}
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

          <div className="wo-tiles">
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

      {expanded && (
        <div className="wo-modal-overlay" onClick={() => setExpanded(null)}>
          <div
            className={`wo-modal ${seMode && modalFields[expanded] ? 'has-sidebar' : ''}`}
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
                <div ref={expandedRef} className="intimation-docx-preview" />
              </div>

              {seMode && modalFields[expanded] && (
                <div className="wo-modal-sidebar">{modalFields[expanded]}</div>
              )}
            </div>

            <div className="wo-modal-foot">
              {actionError && (
                <div className="notice error" style={{ marginRight: 'auto' }}>
                  <IconWarn /> {actionError}
                </div>
              )}
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
        </div>
      )}
    </div>
  )
}

import { useEffect, useMemo, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import { computeWorkAmounts } from '@core/worksAmounts'
import { wrapAgencyAddress } from '@core/workOrderAgreement'
import {
  zoneAbbr,
  financialYearFromDate,
  formatIndianAmount,
  amountInWords
} from '@core/loaSe'
import type { Office } from '../office'
import type { PlaceholderMatch } from '@core/createDocument'
import { pdfToTextLines } from '../pdfToText'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH, normalizeDocxTextboxes } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell } from './Icons'
import type { ExcelTable } from '@core/types'

interface Props {
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
  /** The chosen office. A Zone with no Circle = Superintending Engineer (zonal) office → the LOA format. */
  office: Office
}

/** Fields the SE LOA needs that no source (Works List / notice / L-1) can supply — the user types them. */
interface LoaManualFields {
  adminSanction: string
  period: string
  itemNo: string
  loaDate: string
}

const LOA_MANUAL_DEFAULTS: LoaManualFields = { adminSanction: '', period: '3 Months', itemNo: '', loaDate: '' }

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

/**
 * EMD @ 1.5% is exempted for works reserved for a particular category — the
 * work name carries a "reserved for SC/ST/WLCCS/…" tag. Matches any
 * "reserved for <category>" wording, case-insensitively.
 */
function isEmdExempt(workName: string): boolean {
  return /reserved\s+for\b/i.test(workName)
}

/** Indian financial year for a date (1 April boundary): 2026-07-27 -> "2026-27". */
function indianFinancialYear(d = new Date()): string {
  const y = d.getFullYear()
  const startY = d.getMonth() >= 3 ? y : y - 1
  return `${startY}-${String((startY + 1) % 100).padStart(2, '0')}`
}

/** "18%", " -5 " -> 18, -5. Blank/unparseable -> undefined. */
function parsePct(v: string | undefined): number | undefined {
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
 */
function resolveValue(
  label: string,
  notice: IntimationNotice,
  pdf: TenderEvaluation,
  row: Record<string, string>,
  office?: Office
): string {
  const est = computeWorkAmounts(row)
  const ecv = notice.ecvRupees ?? pdf.ecvRupees ?? est.ecv ?? null
  const tenderPct = pdf.tenderPercentage ?? parsePct(row['Tender Percentage'])
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
  if (PRICE_BID_DATE_LABELS.has(key)) return pdf.noticeDate ?? ''

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
      if (reserved) return 'Rs. 1 ½ Rs.Exempted/-'
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
  return { name, tag: cat ? `(Reserved for ${cat} only)` : '', isReserved: !!cat }
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
  const tenderPct = pdf.tenderPercentage ?? parsePct(row['Tender Percentage'])
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
      return manual.loaDate
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
    case 'circle line': {
      const circle = (row['Circle'] ?? '').trim()
      const cno = (row['CNO'] ?? '').trim()
      if (!circle) return ''
      return cno ? `${circle} Circle-${cno}` : `${circle} Circle`
    }
    case 'nit no':
      return notice.nitNo ?? pdf.noticeNo ?? row['Tender Notice No'] ?? ''
    case 'nit date':
      return pdf.noticeDate ?? ''
    case 'tender id':
      return pdf.tenderId ?? ''
    case 'price bid opening date':
      return pdf.noticeDate ?? ''
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
 * Give Intimation — fills the bundled Intimation format (a .docx mail-merge
 * template with {{placeholders}}) for a Works List work: most fields come from
 * the picked row, while the agency address (and, when present, agency name /
 * NIT No / ECV / contract value) come from the uploaded Online Intimation
 * (portal HTML or LOA PDF) plus the L-1 selection form. Shows a live
 * docx-preview of the filled letter with Word / PDF / Print. (Note Submitted
 * lives on the Agreement & Work Order tab.)
 */
export default function GiveIntimationTab({ tables, onChange, office }: Props) {
  const table = tables[0] ?? null

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

  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [previewPages, setPreviewPages] = useState(0)

  const noticeInputRef = useRef<HTMLInputElement>(null)
  const pdfInputRef = useRef<HTMLInputElement>(null)
  const previewRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  const selectedRow = table && table.rows.length > 0 ? table.rows[Math.min(rowIndex, table.rows.length - 1)] : null
  // The row whose details (Circle, Reservation, …) support the letter. When the
  // uploaded L-1 matched no Works List row, `selectedRow` is just row 0 (a
  // different work), so ignore it — the SE LOA fills those details blank rather
  // than borrowing another work's. A matched L-1 sets rowIndex to that row.
  const detailsRow = worksRowMatched === false ? {} : (selectedRow ?? {})

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

  // Re-fill every placeholder whenever the row / notice / PDF / manual fields change.
  useEffect(() => {
    if (labels.length === 0) return
    setValues(() => {
      const next: Record<string, string> = {}
      for (const label of labels)
        next[label] = seMode
          ? resolveLoaValue(label, notice ?? {}, pdfEval ?? {}, detailsRow, office, manual)
          : resolveValue(label, notice ?? {}, pdfEval ?? {}, detailsRow, office)
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
        const { table: updated, matchedCount, matchedRowIndices } = updateWorksListFromEvaluations(table, [ev], embeddings)
        if (matchedCount > 0) {
          onChange(updated)
          // Select the row that was actually matched — including an embedding
          // (wording-drift) match, which has no exact name to re-derive from.
          // Without this the selection stays on row 0 and the letter fills from
          // a completely different work.
          const idx = matchedRowIndices[0]
          if (idx != null && idx >= 0) setRowIndex(idx)
          setWorksRowMatched(true)
          setPdfStatus(`Updated "${ev.nameOfWork}" in the Works List (Tender ID, Notice No/Date, ECV, Agency, Tender %, Contract Amount).`)
        } else {
          setWorksRowMatched(false)
          setPdfStatus(`Read the PDF, but no Works List row matched "${ev.nameOfWork}" — add this work to the Works List so the Intimation letter fills the correct row.`)
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
  const bothUploaded = !!templateB64 && !!notice && !!pdfEval && !workMismatch

  // Live docx preview of the filled letter — refreshed whenever the values change.
  useEffect(() => {
    if (!bothUploaded) {
      if (previewRef.current) previewRef.current.innerHTML = ''
      setPreviewPages(0)
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
        setPreviewPages(container.querySelectorAll('section.docx').length)
      } catch (e) {
        setActionError(e instanceof Error ? e.message : String(e))
      }
    })()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templateB64, values, bothUploaded])

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

  return (
    <div className="card">
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <div className="empty">
        <IconBell />
        <p>
          Upload the <strong>Online Intimation</strong> and the <strong>L1 selection form</strong> to build the
          intimation letter. The L1 form also updates that work's row in the Works List. The letter previews below
          once both are uploaded.
        </p>
        <div className="boq-actions">
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
        </div>
        {noticeName && <p className="estimate-hint">Address read from {noticeName}</p>}
        {pdfName && <p className="estimate-hint">Tender details read from {pdfName}</p>}
        {pdfStatus && (
          <div className="notice ok">
            <IconBell /> {pdfStatus}
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

      {noWorks ? (
        <div className="notice">Add works to the Works List first — the Intimation letter is filled from a work's row.</div>
      ) : bothUploaded ? (
        <div className="estimate-body">
          {seMode && (
            <div className="loa-manual-fields">
              <span className="estimate-preview-title">
                Superintending Engineer (LOA) — fields to fill in
              </span>
              <div className="loa-manual-grid">
                <label>
                  Admin. Sanction Value
                  <input
                    type="text"
                    placeholder="e.g. 16850000"
                    value={manual.adminSanction}
                    onChange={(e) => setManual((m) => ({ ...m, adminSanction: e.target.value }))}
                  />
                </label>
                <label>
                  Period of Completion
                  <input
                    type="text"
                    placeholder="e.g. 3 Months"
                    value={manual.period}
                    onChange={(e) => setManual((m) => ({ ...m, period: e.target.value }))}
                  />
                </label>
                <label>
                  Item No
                  <input
                    type="text"
                    placeholder="e.g. 01"
                    value={manual.itemNo}
                    onChange={(e) => setManual((m) => ({ ...m, itemNo: e.target.value }))}
                  />
                </label>
                <label>
                  LOA Date
                  <input
                    type="text"
                    placeholder="e.g. 20.03.2026"
                    value={manual.loaDate}
                    onChange={(e) => setManual((m) => ({ ...m, loaDate: e.target.value }))}
                  />
                </label>
              </div>
            </div>
          )}
          <div className="estimate-preview">
            <span className="estimate-preview-title">Live Preview{previewPages > 1 ? ` — ${previewPages} pages` : ''}</span>
            <div className="estimate-preview-scroll">
              <div ref={previewRef} className="intimation-docx-preview" />
            </div>
            <div className="doc-sheet-footer">
              <span className="estimate-hint">Updates live as the details fill in.</span>
              <button className="primary" onClick={() => downloadIntimation(['docx'])} disabled={busy !== null}>
                <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
              </button>
              <button className="primary" onClick={() => downloadIntimation(['pdf'])} disabled={busy !== null}>
                <IconDownload /> {busy === 'pdf' ? 'Saving…' : 'PDF'}
              </button>
              <button className="primary" onClick={printIntimation} disabled={busy !== null}>
                <IconPrint /> {busy === 'print' ? 'Opening…' : 'Print'}
              </button>
            </div>
            {actionError && (
              <div className="notice error">
                <IconWarn /> {actionError}
              </div>
            )}
            {actionSaved && <p className="estimate-hint">{actionSaved}</p>}
          </div>
        </div>
      ) : null}
    </div>
  )
}

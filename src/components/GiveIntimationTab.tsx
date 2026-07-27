import { useEffect, useMemo, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { updateWorksListFromEvaluations } from '@core/worksTenderUpdate'
import { computeWorkAmounts } from '@core/worksAmounts'
import type { PlaceholderMatch } from '@core/createDocument'
import { pdfToTextLines } from '../pdfToText'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell } from './Icons'
import type { ExcelTable } from '@core/types'

interface Props {
  tables: ExcelTable[]
  onChange: (table: ExcelTable) => void
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

function rowLabel(row: Record<string, string>, headers: string[], index: number): string {
  const nameCol = headers.find((h) => /name of (the )?work/i.test(h)) ?? headers[0]
  const v = nameCol ? (row[nameCol] ?? '').trim() : ''
  return v ? `${index + 1}. ${v}` : `Row ${index + 1}`
}

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
 * placeholders share one value (see PRICE_BID_DATE_LABELS). Every value is
 * pre-filled but editable, so a misread field can be fixed before generating.
 */
function resolveValue(
  label: string,
  notice: IntimationNotice,
  pdf: TenderEvaluation,
  row: Record<string, string>
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
  const reserved = isEmdExempt(row['Name of the work'] ?? '')

  const key = norm(label)
  // Both the letter "Date:" and the "price bid opened Date" carry one value,
  // taken from the uploaded L1 selection form. Pre-filled, still editable.
  if (PRICE_BID_DATE_LABELS.has(key)) return pdf.noticeDate ?? ''

  switch (key) {
    case 'agency name':
    case 'name of the agency':
      return notice.agencyName ?? pdf.l1AgencyName ?? row['Name of the Agency'] ?? ''
    case 'address of the agency':
      return notice.address ?? row['Address of the agency'] ?? ''
    case 'agency phone number':
    case 'phone number of the agency':
      return row['Phone number of the agency'] ?? ''
    case 'circle':
      return row['Circle'] ?? ''
    case 'cno':
      return row['CNO'] ?? ''
    case 'zone':
      return row['Zone'] ?? ''
    case 'financial year':
      return indianFinancialYear()
    case 'name of the work':
      return row['Name of the work'] ?? ''
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
      // The whole "(…)" content: reserved -> Exempted; else EMD, plus ASD above 25%.
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

/**
 * Give Intimation — fills the bundled Intimation format (a .docx mail-merge
 * template with {{placeholders}}) for a Works List work: most fields come
 * from the picked Works List row, while the agency's address (and, when
 * present, agency name / NIT No / ECV / contract value) come from an uploaded
 * portal "View Intimation Notice" page (.html). Reuses the same docx
 * placeholder-fill + docx-preview + export/print pipeline as Issue Document.
 */
export default function GiveIntimationTab({ tables, onChange }: Props) {
  const table = tables[0] ?? null

  const [templateB64, setTemplateB64] = useState<string | null>(null)
  const [labels, setLabels] = useState<string[]>([])
  const [loadError, setLoadError] = useState<string | null>(null)

  const [rowIndex, setRowIndex] = useState(0)
  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  const [pdfEval, setPdfEval] = useState<TenderEvaluation | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pdfStatus, setPdfStatus] = useState<string | null>(null)
  const [values, setValues] = useState<Record<string, string>>({})
  // Labels the user has hand-edited — kept out of the auto-fill so a re-pick
  // or re-upload never clobbers a manual correction.
  const [touched, setTouched] = useState<Record<string, boolean>>({})

  const [busy, setBusy] = useState<null | 'preview' | 'download' | 'print' | 'pdf'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [previewPages, setPreviewPages] = useState(0)

  const fileInputRef = useRef<HTMLInputElement>(null)
  const pdfInputRef = useRef<HTMLInputElement>(null)
  const previewRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  const selectedRow = table && table.rows.length > 0 ? table.rows[Math.min(rowIndex, table.rows.length - 1)] : null

  // Load the bundled Intimation format once, and read its placeholders.
  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const b64 = await api.intimationTemplate()
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
  }, [])

  // Re-fill every untouched placeholder whenever the picked row, uploaded
  // notice, or uploaded PDF changes — a hand-edited field (touched) is left
  // exactly as typed.
  useEffect(() => {
    if (labels.length === 0) return
    setValues((prev) => {
      const next = { ...prev }
      for (const label of labels) {
        if (touched[label]) continue
        next[label] = resolveValue(label, notice ?? {}, pdfEval ?? {}, selectedRow ?? {})
      }
      return next
    })
    // selectedRow identity changes with rowIndex/table; notice/pdfEval per upload.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [labels, rowIndex, notice, pdfEval, table])

  function updateValue(label: string, value: string) {
    // The letter date and the "price bid opened" date are one and the same —
    // editing either mirrors to every price-bid-date placeholder in the doc.
    const mirror = PRICE_BID_DATE_LABELS.has(norm(label))
    const targets = mirror ? labels.filter((l) => PRICE_BID_DATE_LABELS.has(norm(l))) : [label]
    setTouched((prev) => {
      const next = { ...prev }
      for (const l of targets) next[l] = true
      return next
    })
    setValues((prev) => {
      const next = { ...prev }
      for (const l of targets) next[l] = value
      return next
    })
  }

  // The Online Intimation can be uploaded as either the portal "View
  // Intimation Notice" .html page or the printed Intimation / Letter of
  // Acceptance .pdf — both carry the agency, address, NIT No, ECV and
  // contract value, so parse by file type.
  async function handleNoticeFile(file: File) {
    setActionError(null)
    try {
      const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
      const parsed = isPdf
        ? parseIntimationNoticeText(await pdfToTextLines(file))
        : parseIntimationNotice(await file.text())
      setNotice(parsed)
      setNoticeName(file.name)
      setTouched({}) // a fresh notice supersedes prior manual edits
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    }
  }

  // A tender-evaluation PDF both fills the letter (NIT/ECV/Tender %/contract/
  // L-1 agency) and updates the Works List row for that work — matched by
  // Name of Work — so the database reflects the award in one step. The
  // matched work is auto-selected so the preview shows it.
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
      setTouched({})

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
        const { table: updated, matchedCount } = updateWorksListFromEvaluations(table, [ev], embeddings)
        if (matchedCount > 0) {
          onChange(updated)
          const nh = updated.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
          const idx = nh
            ? updated.rows.findIndex((r) => norm(r[nh] ?? '') === norm(ev.nameOfWork!))
            : -1
          if (idx >= 0) setRowIndex(idx)
          setPdfStatus(`Updated "${ev.nameOfWork}" in the Works List (Tender ID, Notice No/Date, ECV, Agency, Tender %, Contract Amount).`)
        } else {
          setPdfStatus(`Read the PDF, but no Works List row matched "${ev.nameOfWork}" — the letter is filled, the database wasn't changed.`)
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
    // Fill each {{Label}} directly from its edited value: map every label to
    // itself so fillPlaceholdersInDocument reads values[label] (rather than
    // resolving through a Works List column).
    const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(templateB64, resolved, values)
  }

  async function showPreview() {
    setBusy('preview')
    setActionError(null)
    try {
      const filled = await fillTemplate()
      const container = previewRef.current
      if (!container) return
      container.innerHTML = ''
      await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
      setPreviewPages(container.querySelectorAll('section.docx').length)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function downloadIntimation(formats: ('docx' | 'pdf')[]) {
    setBusy('download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillTemplate()
      const name = `Intimation${selectedRow?.['Name of the Agency'] ? ` - ${selectedRow['Name of the Agency']}` : ''}`
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
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // Guard: the Online Intimation and the L1 selection form must describe the
  // same work (matched by NIT No, or agency name when the NIT No is absent) —
  // otherwise the letter would splice one work's agency onto another's tender.
  const workMatch = useMemo(
    () => (notice && pdfEval ? checkSameWork(notice, pdfEval) : null),
    [notice, pdfEval]
  )
  const workMismatch = workMatch?.status === 'mismatch'

  // Only build the intimation letter once BOTH the Online Intimation (LOA) and
  // the L1 selection form are uploaded, and they belong to the same work — no
  // preview is shown before that.
  const bothUploaded = !!notice && !!pdfEval && !workMismatch

  // Auto-refresh the preview whenever the filled values change — but only after
  // both documents are in; otherwise keep the preview area empty.
  useEffect(() => {
    if (!templateB64 || labels.length === 0) return
    if (!bothUploaded) {
      if (previewRef.current) previewRef.current.innerHTML = ''
      setPreviewPages(0)
      return
    }
    void showPreview()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templateB64, values, bothUploaded])

  const noWorks = !table || table.rows.length === 0

  return (
    <div className="card">
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <div className="empty">
        <IconBell />
        <p>
          Upload the <strong>Online Intimation</strong> and the <strong>L1 selection form</strong> to build the
          intimation letter. The L1 form also updates that work's row in the Works List. The details and preview
          appear once both are uploaded.
        </p>
        <div className="boq-actions">
          <button className="primary upload-btn" onClick={() => fileInputRef.current?.click()} disabled={!templateB64}>
            <IconFolder /> {notice ? 'Change Online Intimation' : 'Upload Online Intimation'}
          </button>
          <button className="primary upload-btn" onClick={() => pdfInputRef.current?.click()} disabled={!templateB64 || busy === 'pdf'}>
            <IconFolder /> {busy === 'pdf' ? 'Reading PDF…' : pdfEval ? 'Change L1 selection form' : 'Upload L1 selection form'}
          </button>
          <input
            ref={fileInputRef}
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

      {noWorks ? (
        <div className="notice">Add works to the Works List first — the Intimation letter is filled from a work's row.</div>
      ) : bothUploaded ? (
        <div className="estimate-body">
          <div className="estimate-preview">
              <span className="estimate-preview-title">Live Preview{previewPages > 1 ? ` — ${previewPages} pages` : ''}</span>
              <div className="estimate-preview-scroll">
                <div ref={previewRef} className="intimation-docx-preview" />
              </div>
              <div className="doc-sheet-footer">
                <span className="estimate-hint">Updates live as you fill in the details.</span>
                <button className="primary" onClick={() => downloadIntimation(['docx'])} disabled={busy !== null}>
                  <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
                </button>
                <button className="primary" onClick={() => downloadIntimation(['pdf'])} disabled={busy !== null}>
                  <IconDownload /> {busy === 'download' ? 'Saving…' : 'PDF'}
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

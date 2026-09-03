import { useEffect, useMemo, useRef, useState } from 'react'
import { renderAsync } from '../lazyDocxPreview'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { circleFromNit } from '@core/workOrderAgreement'
import { resolveIntimationValue } from '@core/intimationFill'
import { CMC_ZONE_CIRCLES, resolveFromDirectory, corporationByName } from '../zoneCircleDirectory'
import { officeScopedKey, TEMPLATE_KEYS, type Office } from '../office'
import type { PlaceholderMatch } from '@core/createDocument'
import { pdfToTextLines } from '../pdfToText'
import { base64ToUint8, PAGE_WIDTH, renderDocPreview, DOCX_PREVIEW_OPTIONS, normalizeDocxTextboxes } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell } from './Icons'

interface Props {
  /** The chosen office — only its Corporation decides which zone/circle directory
   * to resolve the work name against; the circle itself comes from the work name,
   * so this tool works for any circle regardless of which office is selected. */
  office?: Office
}

/** The office details this tool stamps on the letter — resolved from the work name. */
interface DetectedOffice {
  circle: string
  cno: string
  zone: string
  /** How the circle was found — for the "detected from…" hint. */
  source: 'work-name' | 'nit' | 'manual' | 'none'
}

/**
 * Resolve which circle's office details the Intimation should carry, purely from
 * the uploaded L-1 — no Works List, no selected circle. The work name almost
 * always names its circle ("…in Nizampet circle…"), so that's matched first
 * (resolveFromDirectory); only if the name carries no circle do we fall back to
 * the "/EE/<Circle> Circle-NN" tag in the NIT No. A manual override always wins.
 */
function detectOffice(
  entries: typeof CMC_ZONE_CIRCLES,
  ev: TenderEvaluation | null,
  notice: IntimationNotice | null,
  override: string
): DetectedOffice {
  if (override) {
    const e = entries.find((c) => c.circle === override)
    return { circle: override, cno: e?.cno ?? '', zone: e?.zone ?? '', source: 'manual' }
  }
  const byName = resolveFromDirectory(ev?.nameOfWork, entries)
  if (byName.circle) {
    return { circle: byName.circle, cno: byName.cno ?? '', zone: byName.zone ?? '', source: 'work-name' }
  }
  const nit = circleFromNit(ev?.noticeNo || notice?.nitNo)
  if (nit.circle) {
    const e = entries.find((c) => c.circle.toLowerCase() === nit.circle.toLowerCase())
    return { circle: nit.circle, cno: nit.cno || e?.cno || '', zone: e?.zone ?? byName.zone ?? '', source: 'nit' }
  }
  // Nothing matched a known circle — keep any zone the name did resolve to.
  return { circle: '', cno: '', zone: byName.zone ?? '', source: 'none' }
}

/**
 * Tools workspace — Intimation. Fills the bundled EE Intimation format from an
 * uploaded L-1 selection form (+ Online Intimation for the postal address),
 * standing entirely on its own: it works for ANY circle by reading the circle
 * out of the L-1's Name of Work (the office details of that circle are then
 * stamped on the letter), so no Works List row and no matching office login are
 * needed. Mirrors Give Intimation's fill/preview, minus the Works List coupling.
 */
export default function IntimationToolTab({ office }: Props) {
  // Resolve the work name against the chosen corporation's directory, defaulting
  // to CMC (the only fully-populated one) so the tool still works before an
  // office is picked.
  const entries = useMemo(
    () => corporationByName(office?.corporation)?.entries?.length
      ? corporationByName(office?.corporation)!.entries
      : CMC_ZONE_CIRCLES,
    [office?.corporation]
  )
  const corporationName = corporationByName(office?.corporation)?.entries?.length
    ? office!.corporation!
    : 'CMC'

  const [templateB64, setTemplateB64] = useState<string | null>(null)
  const [labels, setLabels] = useState<string[]>([])
  const [loadError, setLoadError] = useState<string | null>(null)

  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  const [pdfEval, setPdfEval] = useState<TenderEvaluation | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pdfStatus, setPdfStatus] = useState<string | null>(null)
  // A hand-picked circle override (the auto-detected one is used until then).
  const [circleOverride, setCircleOverride] = useState('')

  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [previewPages, setPreviewPages] = useState(0)

  const noticeInputRef = useRef<HTMLInputElement>(null)
  const pdfInputRef = useRef<HTMLInputElement>(null)
  const previewRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  const detected = useMemo(
    () => detectOffice(entries, pdfEval, notice, circleOverride),
    [entries, pdfEval, notice, circleOverride]
  )

  // The synthetic Works-List-shaped row: Circle/CNO/Zone from the work name (the
  // whole point of this tool), the rest read from the uploaded L-1 / Intimation
  // inside resolveIntimationValue.
  const row = useMemo<Record<string, string>>(
    () => ({
      Circle: detected.circle,
      CNO: detected.cno,
      Zone: detected.zone,
      Corporation: corporationName,
      'Corp Full': corporationByName(corporationName)?.fullName ?? '',
      'Name of the work': pdfEval?.nameOfWork ?? '',
      'Tender Notice No': pdfEval?.noticeNo || notice?.nitNo || ''
    }),
    [detected, pdfEval, notice, corporationName]
  )

  const values = useMemo<Record<string, string>>(() => {
    const next: Record<string, string> = {}
    for (const label of labels) next[label] = resolveIntimationValue(label, notice ?? {}, pdfEval ?? {}, row)
    return next
  }, [labels, notice, pdfEval, row])

  // Load the EE Intimation format and read its placeholders (once).
  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const intimationVariant = localStorage.getItem(officeScopedKey(TEMPLATE_KEYS.intimation, office)) ?? undefined
        const b64 = await api.intimationTemplate(intimationVariant)
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
  }, [office])

  // The Online Intimation supplies the agency's postal address (and, when the
  // L-1 is missing them, the NIT No / ECV / contract value). Optional here — the
  // L-1 alone carries the work name (→ circle), agency name and amounts.
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

  // The L-1 selection form is the source of truth: its Name of Work resolves the
  // circle, and it carries the agency name and amounts.
  async function handlePdfFile(file: File) {
    setBusy('pdf')
    setActionError(null)
    setPdfStatus(null)
    setCircleOverride('')
    try {
      const lines = await pdfToTextLines(file)
      const ev = parseTenderEvaluation(lines)
      if (!ev.nameOfWork && !ev.tenderId) {
        throw new Error("Couldn't read tender details from that PDF — is it the Commercial Evaluation / Stage Selected page?")
      }
      setPdfEval(ev)
      setPdfName(file.name)
      const found = detectOffice(entries, ev, notice, '')
      setPdfStatus(
        found.circle
          ? `Detected ${found.circle} Circle${found.cno ? `-${found.cno}` : ''}${found.zone ? `, ${found.zone} zone` : ''} from ${found.source === 'nit' ? 'the NIT No' : 'the Name of Work'}.`
          : `Couldn't find a known circle in "${ev.nameOfWork ?? 'the work name'}" — pick the circle below.`
      )
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

  // The Online Intimation and L-1, when both present, must describe the same work.
  const workMatch = useMemo(() => (notice && pdfEval ? checkSameWork(notice, pdfEval) : null), [notice, pdfEval])
  const workMismatch = workMatch?.status === 'mismatch'
  // Ready to preview once the L-1 is in (the Online Intimation only adds the
  // address) and the two uploads — if both present — agree.
  const ready = !!templateB64 && !!pdfEval && !workMismatch

  // Live docx preview of the filled letter — refreshed whenever the values
  // change, so this deliberately stays on the fast, approximate
  // docx-preview.js render rather than the accurate LibreOffice-backed one
  // (printIntimation, below, is where the accurate render belongs): a
  // LibreOffice round trip on every keystroke would make typing feel
  // sluggish.
  useEffect(() => {
    if (!ready) {
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
  }, [templateB64, values, ready])

  async function downloadIntimation(formats: ('docx' | 'pdf')[]) {
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillTemplate()
      const agencyName = notice?.agencyName ?? pdfEval?.l1AgencyName
      const name = `Intimation${agencyName ? ` - ${agencyName}` : ''}`
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
      await renderDocPreview(base64ToUint8(filled), container)
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  return (
    <div className="card">
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <div className="intimation-tool-uploads">
        <div className="boq-actions">
          <button
            className="primary upload-btn"
            onClick={() => pdfInputRef.current?.click()}
            disabled={!templateB64 || busy === 'pdf'}
          >
            <IconFolder /> {busy === 'pdf' ? 'Reading PDF…' : pdfEval ? 'Change L1 selection form' : 'Upload L1 selection form'}
          </button>
          <button className="primary upload-btn" onClick={() => noticeInputRef.current?.click()} disabled={!templateB64}>
            <IconFolder /> {notice ? 'Change Online Intimation' : 'Upload Online Intimation (address)'}
          </button>
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
        </div>
        {pdfName && <p className="estimate-hint">Work &amp; tender details read from {pdfName}</p>}
        {noticeName && <p className="estimate-hint">Address read from {noticeName}</p>}
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

      {pdfEval && (
        <div className="loa-manual-fields">
          <span className="estimate-preview-title">Office (circle from the work name)</span>
          <div className="loa-manual-grid">
            <label>
              Circle
              <select value={detected.circle} onChange={(e) => setCircleOverride(e.target.value)}>
                <option value="">— none detected —</option>
                {entries.map((c) => (
                  <option key={c.circle} value={c.circle}>
                    {c.circle} — {c.cno} ({c.zone})
                  </option>
                ))}
              </select>
            </label>
            <label>
              CNO
              <input type="text" value={detected.cno} readOnly />
            </label>
            <label>
              Zone
              <input type="text" value={detected.zone} readOnly />
            </label>
          </div>
        </div>
      )}

      {ready ? (
        <div className="estimate-body">
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

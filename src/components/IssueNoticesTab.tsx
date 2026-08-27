import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { circleFromNit } from '@core/workOrderAgreement'
import { issueNoticePlaceholders, EMPTY_ISSUE_NOTICE_MANUAL_FIELDS, type IssueNoticeManualFields } from '@core/issueNotice'
import { CMC_ZONE_CIRCLES, resolveFromDirectory, corporationByName } from '../zoneCircleDirectory'
import { isZoneOnlyOffice, type Office } from '../office'
import type { PlaceholderMatch } from '@core/createDocument'
import { pdfToTextLines } from '../pdfToText'
import { base64ToUint8, PAGE_WIDTH, renderDocPreview } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell } from './Icons'
import type { ThemeId } from '../theme'

interface Props {
  /** The chosen office — decides which format this Notice prints as: a
   * Zone-level (SE) office (a Zone picked, no Circle) gets the SE-signed
   * format, anything else (a Circle picked, or no office chosen yet) gets
   * the EE-signed one. Its Corporation also picks which zone/circle
   * directory to resolve the uploaded work's own office block from. */
  office?: Office
  /** Issue Documents tile style, set in Settings → Themes (see theme.ts) — applied to this page's output tiles too. */
  theme: ThemeId
}

/** The office block this Notice's letterhead carries — resolved from the work name. */
interface DetectedOffice {
  circle: string
  cno: string
  zone: string
  /** How it was found — for the "detected from…" hint. */
  source: 'work-name' | 'nit' | 'manual' | 'none'
}

/**
 * Resolve which office block the Notice's letterhead names, purely from the
 * uploaded L-1 — no Works List, no matching login. The work name almost
 * always names its circle ("…in Nizampet circle…"), which the directory maps
 * to a zone (resolveFromDirectory); only if the name carries no circle do we
 * fall back to the "/EE/<Circle> Circle-NN" tag in the NIT No. A manual
 * override (whichever the current office type's picker below sets) always
 * wins — a Circle name in EE mode, resolved via the directory for its zone
 * too; a bare Zone name in SE mode, which has no single circle to go with it.
 */
function detectOffice(
  entries: typeof CMC_ZONE_CIRCLES,
  ev: TenderEvaluation | null,
  notice: IntimationNotice | null,
  override: { circle?: string; zone?: string }
): DetectedOffice {
  if (override.circle) {
    const e = entries.find((c) => c.circle === override.circle)
    return { circle: override.circle, cno: e?.cno ?? '', zone: e?.zone ?? '', source: 'manual' }
  }
  if (override.zone) return { circle: '', cno: '', zone: override.zone, source: 'manual' }
  const byName = resolveFromDirectory(ev?.nameOfWork, entries)
  if (byName.circle) {
    return { circle: byName.circle, cno: byName.cno ?? '', zone: byName.zone ?? '', source: 'work-name' }
  }
  const nit = circleFromNit(ev?.noticeNo || notice?.nitNo)
  if (nit.circle) {
    const e = entries.find((c) => c.circle.toLowerCase() === nit.circle.toLowerCase())
    if (e) return { circle: e.circle, cno: nit.cno || e.cno, zone: e.zone, source: 'nit' }
  }
  // Nothing matched a known circle — keep any zone the name did resolve to.
  return { circle: '', cno: '', zone: byName.zone ?? '', source: 'none' }
}

/**
 * Issue Notices — fills the bundled "Notice for Non-conclusion of Agreement"
 * format from an uploaded L-1 selection form (+ Online Intimation for the
 * agency's address), standing entirely on its own like the Tools Intimation
 * tile: it works for ANY zone/circle by reading the circle out of the L-1's
 * Name of Work, so no Works List row and no matching office login are
 * needed. Which of the two bundled formats is filled — Zone-level (SE) or
 * Circle-level (EE), each with its own letterhead, numbering and signature —
 * follows the office chosen in the sidebar (isZoneOnlyOffice), so an EE
 * office always issues an EE-signed Notice and an SE office an SE-signed one.
 * The LOA No/Date and the agency phone aren't printed on either upload, so
 * they're always hand-typed below.
 */
export default function IssueNoticesTab({ office, theme }: Props) {
  const entries = useMemo(
    () => (corporationByName(office?.corporation)?.entries?.length
      ? corporationByName(office?.corporation)!.entries
      : CMC_ZONE_CIRCLES),
    [office?.corporation]
  )
  const corporationName = corporationByName(office?.corporation)?.entries?.length
    ? office!.corporation!
    : 'CMC'
  const corporationFullName = corporationByName(corporationName)?.fullName
  // Which bundled format this Notice fills — Circle-level (EE) unless the
  // chosen office is unambiguously Zone-level (SE), matching how every other
  // paired EE/SE document in this app decides (WorkOrderAgreementTab's
  // zoneLogin, SE Bid Document, etc.): default to EE, SE only opts in.
  const ee = !isZoneOnlyOffice(office)

  const [templateB64, setTemplateB64] = useState<string | null>(null)
  const [labels, setLabels] = useState<string[]>([])
  const [loadError, setLoadError] = useState<string | null>(null)

  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  const [pdfEval, setPdfEval] = useState<TenderEvaluation | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pdfStatus, setPdfStatus] = useState<string | null>(null)
  // A hand-picked place-name override (the auto-detected circle/zone is used
  // until then) — a Circle name in EE mode, a Zone name in SE mode. Reset
  // whenever the office switches mode, so a leftover Circle pick from EE
  // mode can't silently masquerade as a Zone (or vice versa) in the other.
  const [placeOverride, setPlaceOverride] = useState<{ circle?: string; zone?: string }>({})
  useEffect(() => {
    setPlaceOverride({})
  }, [ee])
  const [manual, setManual] = useState<IssueNoticeManualFields>(EMPTY_ISSUE_NOTICE_MANUAL_FIELDS)

  const [busy, setBusy] = useState<null | 'download' | 'print' | 'pdf'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [previewPages, setPreviewPages] = useState(0)
  // Whether the tile's full-size preview modal is open — matches the
  // Work Order/Agreement Bond tile catalog's own "click a tile, see the
  // accurate full-size render + a field sidebar" pattern.
  const [expanded, setExpanded] = useState(false)
  const [tileFailed, setTileFailed] = useState(false)

  const noticeInputRef = useRef<HTMLInputElement>(null)
  const pdfInputRef = useRef<HTMLInputElement>(null)
  const tileRef = useRef<HTMLDivElement>(null)
  const expandedRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  const detected = useMemo(
    () => detectOffice(entries, pdfEval, notice, placeOverride),
    [entries, pdfEval, notice, placeOverride]
  )

  const values = useMemo<Record<string, string>>(
    () =>
      issueNoticePlaceholders(
        notice ?? {},
        pdfEval ?? {},
        manual,
        ee
          ? { circle: detected.circle, cno: detected.cno, corporation: corporationName, corporationFullName }
          : { zone: detected.zone, corporation: corporationName, corporationFullName }
      ),
    [notice, pdfEval, manual, ee, detected.circle, detected.cno, detected.zone, corporationName, corporationFullName]
  )

  function setManualField<K extends keyof IssueNoticeManualFields>(key: K, v: string) {
    setManual((prev) => ({ ...prev, [key]: v }))
  }

  // Load whichever format the office needs (and its placeholders) — re-fetches
  // if the office switches between EE and SE mid-session.
  useEffect(() => {
    let cancelled = false
    setTemplateB64(null)
    void (async () => {
      try {
        const b64 = await api.issueNoticeTemplate(ee)
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
  }, [ee])

  // The Online Intimation supplies the agency's postal address (and, when the
  // L-1 is missing them, the NIT No / ECV / contract value). Optional here — the
  // L-1 alone carries the work name (→ office block), agency name and amounts.
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
  // office block, and it carries the agency name and amounts.
  async function handlePdfFile(file: File) {
    setBusy('pdf')
    setActionError(null)
    setPdfStatus(null)
    setPlaceOverride({})
    try {
      const lines = await pdfToTextLines(file)
      const ev = parseTenderEvaluation(lines)
      if (!ev.nameOfWork && !ev.tenderId) {
        throw new Error("Couldn't read tender details from that PDF — is it the Commercial Evaluation / Stage Selected page?")
      }
      setPdfEval(ev)
      setPdfName(file.name)
      const found = detectOffice(entries, ev, notice, {})
      const label = ee
        ? found.circle
          ? `${found.circle} Circle${found.cno ? `-${found.cno}` : ''}`
          : ''
        : found.zone
      setPdfStatus(
        label
          ? `Detected ${label} from ${found.source === 'nit' ? 'the NIT No' : 'the Name of Work'}.`
          : `Couldn't find a known ${ee ? 'circle' : 'zone'} in "${ev.nameOfWork ?? 'the work name'}" — pick it below.`
      )
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function fillTemplate(): Promise<string> {
    if (!templateB64) throw new Error('Notice format not loaded yet.')
    const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(templateB64, resolved, values)
  }

  // The Online Intimation and L-1, when both present, must describe the same work.
  const workMatch = useMemo(() => (notice && pdfEval ? checkSameWork(notice, pdfEval) : null), [notice, pdfEval])
  const workMismatch = workMatch?.status === 'mismatch'
  // Ready to preview once the L-1 is in (the Online Intimation only adds the
  // address) and the two uploads — if both present — agree.
  const ready = !!templateB64 && !!pdfEval && !workMismatch

  // The tile's live thumbnail — the same accurate, LibreOffice-backed render
  // the expanded modal uses (renderDocPreview, below), not the fast
  // docx-preview.js one: with only one tile here (unlike the Work Order/
  // Agreement catalog's 11), a single accurate conversion is affordable.
  // Debounced — `values` changes on every keystroke in the LOA/Notice fields
  // — so typing doesn't fire a LibreOffice conversion per character.
  useEffect(() => {
    if (!ready) {
      if (tileRef.current) tileRef.current.innerHTML = ''
      setTileFailed(false)
      return
    }
    const timer = setTimeout(() => {
      const container = tileRef.current
      if (!container) return
      void fillTemplate()
        .then((filled) => renderDocPreview(base64ToUint8(filled), container))
        .then(() => setTileFailed(false))
        .catch(() => setTileFailed(true))
    }, 600)
    return () => clearTimeout(timer)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templateB64, values, ready])

  // The full-size, same accurate render — re-rendered into the modal
  // whenever it's open and the values change, so the entire document (every
  // page) is readable, not just the cropped tile thumbnail.
  useEffect(() => {
    if (!expanded || !ready) return
    const container = expandedRef.current
    if (!container) return
    setActionError(null)
    void fillTemplate()
      .then((filled) => renderDocPreview(base64ToUint8(filled), container))
      .then(({ pageCount }) => setPreviewPages(pageCount))
      .catch((e) => setActionError(e instanceof Error ? e.message : String(e)))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [expanded, templateB64, values, ready])

  async function downloadNotice(formats: ('docx' | 'pdf')[]) {
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillTemplate()
      const agencyName = notice?.agencyName ?? pdfEval?.l1AgencyName
      const name = `Notice${agencyName ? ` - ${agencyName}` : ''}`
      const res = await api.exportCreatedDocument(filled, name, formats)
      setActionSaved(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function printNotice() {
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
        <div className="boq-actions boq-actions--start">
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
        <p className="estimate-hint">
          {ee ? 'Executive Engineer format' : 'Superintending Engineer format'} — follows the office picked on the Works List page.
        </p>
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
          <span className="estimate-preview-title">Office ({ee ? 'circle' : 'zone'} from the work name)</span>
          <div className="loa-manual-grid">
            {ee ? (
              <label>
                Circle
                <select value={detected.circle} onChange={(e) => setPlaceOverride({ circle: e.target.value })}>
                  <option value="">— none detected —</option>
                  {entries.map((c) => (
                    <option key={c.circle} value={c.circle}>
                      {c.circle} — {c.cno} ({c.zone})
                    </option>
                  ))}
                </select>
              </label>
            ) : (
              <label>
                Zone
                <select value={detected.zone} onChange={(e) => setPlaceOverride({ zone: e.target.value })}>
                  <option value="">— none detected —</option>
                  {[...new Set(entries.map((c) => c.zone))].map((z) => (
                    <option key={z} value={z}>
                      {z}
                    </option>
                  ))}
                </select>
              </label>
            )}
          </div>
        </div>
      )}

      {ready && (
        <div className={`wo-tiles${theme === 'flat1' ? ' wo-tiles-flat' : ''}`}>
          <div
            role="button"
            tabIndex={0}
            className="wo-tile"
            onClick={() => setExpanded(true)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault()
                setExpanded(true)
              }
            }}
            title="Notice for Non-conclusion of Agreement"
          >
            <div className="wo-tile-preview">
              {tileFailed ? (
                <div className="settings-tile-error">
                  <IconWarn /> Couldn't load
                </div>
              ) : (
                <div ref={tileRef} className="wo-tile-doc" />
              )}
              <span className="wo-tile-open">Click to preview</span>
            </div>
            <div className="wo-tile-foot">Notice for Non-conclusion of Agreement</div>
          </div>
        </div>
      )}

      {expanded &&
        createPortal(
          <div className="wo-modal-overlay" onClick={() => setExpanded(false)}>
            <div className="wo-modal has-sidebar" onClick={(e) => e.stopPropagation()}>
              <div className="wo-modal-head">
                <span className="wo-modal-title">
                  Notice for Non-conclusion of Agreement{previewPages > 1 ? ` — ${previewPages} pages` : ''}
                </span>
                <button className="wo-modal-close" onClick={() => setExpanded(false)} title="Close" aria-label="Close">
                  ×
                </button>
              </div>
              <div className="wo-modal-main">
                <div className="wo-modal-body">
                  <div ref={expandedRef} className="intimation-docx-preview" />
                </div>
                <div className="wo-modal-sidebar">
                  <span className="estimate-preview-title">LOA details (not on either upload — type these in)</span>
                  <label className="wo-date-field">
                    <span>LOA No</span>
                    <input
                      type="text"
                      value={manual.loaNo}
                      onChange={(e) => setManualField('loaNo', e.target.value)}
                      placeholder="e.g. 27"
                    />
                  </label>
                  <label className="wo-date-field">
                    <span>LOA Date</span>
                    <input
                      type="text"
                      value={manual.loaDate}
                      onChange={(e) => setManualField('loaDate', e.target.value)}
                      placeholder="dd.mm.yyyy"
                    />
                  </label>
                  <label className="wo-date-field">
                    <span>Notice Date</span>
                    <input
                      type="text"
                      value={manual.noticeDate}
                      onChange={(e) => setManualField('noticeDate', e.target.value)}
                      placeholder="dd.mm.yyyy"
                    />
                  </label>
                  <label className="wo-date-field">
                    <span>Agency Phone (optional)</span>
                    <input
                      type="tel"
                      value={manual.agencyPhone}
                      onChange={(e) => setManualField('agencyPhone', e.target.value)}
                      placeholder="10-digit mobile number"
                    />
                  </label>
                </div>
              </div>
              <div className="wo-modal-foot">
                {actionError && (
                  <div className="notice error" style={{ marginRight: 'auto' }}>
                    <IconWarn /> {actionError}
                  </div>
                )}
                {actionSaved && (
                  <span className="estimate-hint" style={{ marginRight: 'auto' }}>
                    {actionSaved}
                  </span>
                )}
                <button className="primary" onClick={() => downloadNotice(['docx'])} disabled={busy !== null}>
                  <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
                </button>
                <button className="primary" onClick={() => downloadNotice(['pdf'])} disabled={busy !== null}>
                  <IconDownload /> {busy === 'pdf' ? 'Saving…' : 'PDF'}
                </button>
                <button className="primary" onClick={printNotice} disabled={busy !== null}>
                  <IconPrint /> {busy === 'print' ? 'Opening…' : 'Print'}
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
    </div>
  )
}

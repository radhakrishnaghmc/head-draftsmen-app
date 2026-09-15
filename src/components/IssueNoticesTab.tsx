import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { circleFromNit } from '@core/workOrderAgreement'
import {
  issueNoticePlaceholders,
  EMPTY_ISSUE_NOTICE_MANUAL_FIELDS,
  type IssueNoticeManualFields,
  type IssueNoticeOffice
} from '@core/issueNotice'
import type { PendingNoticeItem } from '@core/types'
import { CMC_ZONE_CIRCLES, resolveFromDirectory, corporationByName } from '../zoneCircleDirectory'
import { isZoneOnlyOffice, officeScopedKey, TEMPLATE_KEYS, TEMPLATE_VARIANT_CHANGED_EVENT, type Office } from '../office'
import type { PlaceholderMatch } from '@core/createDocument'
import { pdfToTextLinesOrOcr } from '../pdfToText'
import { base64ToUint8, PAGE_WIDTH, renderDocPreview } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell, IconTrash, IconCheck, IconRestore } from './Icons'

interface Props {
  /** chosen office — decides which bundled format the manual-upload path
   * defaults to (Zone-level SE vs. Circle-level EE), same rule every tile
   * uses for the work it belongs to. */
  office?: Office
  /** This office's Notice tiles — one per work whose Intimation/LOA has been
   * issued (auto-added by Give Intimation) or manually uploaded here. */
  items: PendingNoticeItem[]
  onChange: (next: PendingNoticeItem[]) => void
}

/** office block a Notice's letterhead carries — resolved from the work name. */
interface DetectedOffice {
  circle: string
  cno: string
  zone: string
  /** How it was found — for the "detected from…" hint. */
  source: 'work-name' | 'nit' | 'manual' | 'none'
}

/**
 * Resolve the office block a Notice's letterhead names, from an uploaded L-1
 * selection form's own Name of Work (falls back to the NIT No's own
 * "/EE/<Circle> Circle-NN" tail) — same resolution the Tools-tab Intimation
 * and Work Order tabs use. Looks up the EE circle for its zone too; bare
 * Zone name only, in SE mode, since no single circle goes with it.
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
  // Nothing matched a known circle — keep any zone name it did resolve to.
  return { circle: '', cno: '', zone: byName.zone ?? '', source: 'none' }
}

/** Builds the placeholder map for a stored `PendingNoticeItem` — same
 * `issueNoticePlaceholders` the manual-upload path uses, fed from the
 * item's own flattened fields instead of two freshly-parsed uploads.
 * `Item No` isn't parseable off a plain work name (no "item N" tag on
 * it once it's been flattened into the item), so it's overridden from
 * the item's own `itemNo` afterward, when present (SE offices only). */
function placeholdersForItem(item: PendingNoticeItem): Record<string, string> {
  const notice: IntimationNotice = {
    agencyName: item.agencyName,
    address: item.address,
    nitNo: item.nitNo,
    nitDate: item.nitDate,
    ecvRupees: item.ecvRupees,
    contractRupees: item.contractRupees
  }
  const pdf: TenderEvaluation = {
    nameOfWork: item.workName,
    tenderPercentage: item.tenderPercentage
  }
  const corporationFullName = corporationByName(item.corporation)?.fullName
  const office: IssueNoticeOffice = item.seMode
    ? { zone: item.zone, corporation: item.corporation, corporationFullName }
    : { circle: item.circle, cno: item.cno, corporation: item.corporation, corporationFullName }
  const values = issueNoticePlaceholders(notice, pdf, item.manual, office)
  if (item.itemNo) values['Item No'] = item.itemNo
  return values
}

interface NoticeTileProps {
  item: PendingNoticeItem
  templateB64: string | null
  labels: string[]
  onSave: (patch: Partial<PendingNoticeItem>) => void
  onRemove: () => void
  onRestore: () => void
}

function NoticeTile({ item, templateB64, labels, onSave, onRemove, onRestore }: NoticeTileProps) {
  const [expanded, setExpanded] = useState(false)
  const [busy, setBusy] = useState<'download' | 'pdf' | 'print' | null>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [previewPages, setPreviewPages] = useState(1)
  const [tileFailed, setTileFailed] = useState(false)
  const tileRef = useRef<HTMLDivElement>(null)
  const expandedRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  const values = useMemo(() => placeholdersForItem(item), [item])
  const ready = !!templateB64 && labels.length > 0

  async function fillTemplate(): Promise<string> {
    if (!templateB64) throw new Error('Notice format not loaded yet.')
    const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(templateB64, resolved, values)
  }

  // Tile's live thumbnail — same accurate, LibreOffice-backed preview every
  // other document tile in the app uses.
  useEffect(() => {
    if (!ready) return
    const timer = setTimeout(() => {
      const container = tileRef.current
      if (!container) return
      void fillTemplate()
        .then((filled) => renderDocPreview(base64ToUint8(filled), container))
        .then(() => setTileFailed(false))
        .catch(() => setTileFailed(true))
    }, 300)
    return () => clearTimeout(timer)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templateB64, values, ready])

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

  function setManualField<K extends keyof IssueNoticeManualFields>(key: K, v: string) {
    onSave({ manual: { ...item.manual, [key]: v } })
  }

  async function downloadNotice(formats: ('docx' | 'pdf')[]) {
    setBusy(formats[0] === 'pdf' ? 'pdf' : 'download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillTemplate()
      const nitNo = item.nitNo?.replace(/[/\\]/g, '-')
      const name =
        nitNo && item.agencyName
          ? `${nitNo} - ${item.agencyName}`
          : `Notice${item.agencyName ? ` - ${item.agencyName}` : nitNo ? ` - ${nitNo}` : ''}`
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
    <>
      <div
        ref={printScratchRef}
        style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }}
        aria-hidden
      />
      <div className="wo-tile" role="button" tabIndex={0} onClick={() => setExpanded(true)}
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
        <div className="wo-tile-foot">
          {item.agencyName || item.workName || 'Notice for Non-conclusion of Agreement'}
          {item.done && ' · Resolved'}
        </div>
        <div className="wo-tile-actions" onClick={(e) => e.stopPropagation()}>
          {item.done ? (
            <button className="ghost small" onClick={onRestore} title="Mark active again">
              <IconRestore /> Reopen
            </button>
          ) : (
            <button className="ghost small" onClick={() => onSave({ done: true })} title="Agreement concluded — mark resolved">
              <IconCheck /> Resolved
            </button>
          )}
          <button className="ghost small" onClick={onRemove} title="Remove">
            <IconTrash />
          </button>
        </div>
      </div>

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
                  <span className="estimate-preview-title">LOA details (hand-recorded — type in)</span>
                  <label className="wo-date-field">
                    <span>LOA No</span>
                    <input
                      type="text"
                      value={item.manual.loaNo}
                      onChange={(e) => setManualField('loaNo', e.target.value)}
                      placeholder="e.g. 27"
                    />
                  </label>
                  <label className="wo-date-field">
                    <span>LOA Date</span>
                    <input
                      type="text"
                      value={item.manual.loaDate}
                      onChange={(e) => setManualField('loaDate', e.target.value)}
                      placeholder="dd.mm.yyyy"
                    />
                  </label>
                  <label className="wo-date-field">
                    <span>Notice Date</span>
                    <input
                      type="text"
                      value={item.manual.noticeDate}
                      onChange={(e) => setManualField('noticeDate', e.target.value)}
                      placeholder="dd.mm.yyyy"
                    />
                  </label>
                  <label className="wo-date-field">
                    <span>Agency phone (optional)</span>
                    <input
                      type="text"
                      value={item.manual.agencyPhone}
                      onChange={(e) => setManualField('agencyPhone', e.target.value)}
                      placeholder="10-digit"
                    />
                  </label>
                </div>
              </div>
              <div className="wo-modal-foot">
                {actionError && (
                  <div className="notice error">
                    <IconWarn /> {actionError}
                  </div>
                )}
                {actionSaved && <span className="estimate-hint">{actionSaved}</span>}
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
    </>
  )
}

/**
 * Issue Notices — one tile per work whose Intimation/LOA has been issued.
 * Tiles are usually prepared automatically the moment Give Intimation
 * issues an Intimation/LOA (7 days' grace, notice on the 8th working day —
 * see App.tsx's handleIntimationIssued); the upload flow below stays as a
 * manual fallback (e.g. a Notice needed for a work issued before this
 * existed) and simply appends another tile the same way.
 */
export default function IssueNoticesTab({ office, items, onChange }: Props) {
  const entries = useMemo(
    () =>
      corporationByName(office?.corporation)?.entries?.length
        ? corporationByName(office?.corporation)!.entries
        : CMC_ZONE_CIRCLES,
    [office?.corporation]
  )
  const corporationName = corporationByName(office?.corporation)?.entries?.length ? office!.corporation! : 'CMC'
  const ee = !isZoneOnlyOffice(office)

  // Both bundled formats (SE/EE) are fetched once, up front — tiles can be a
  // mix of seMode (works from different offices can be mid-flight at once).
  type TemplateSlot = { b64: string | null; labels: string[] }
  const EMPTY_SLOT: TemplateSlot = { b64: null, labels: [] }
  const [templates, setTemplates] = useState<Record<string, TemplateSlot>>({
    se: EMPTY_SLOT,
    ee: EMPTY_SLOT
  })
  const [loadError, setLoadError] = useState<string | null>(null)

  const noticeVariantKey = officeScopedKey(TEMPLATE_KEYS.notice, office)

  useEffect(() => {
    let cancelled = false

    function load() {
      void (async () => {
        try {
          const keys: [string, boolean][] = [
            ['se', false],
            ['ee', true]
          ]
          const version = localStorage.getItem(noticeVariantKey) === '2' ? 2 : 1
          const slots = await Promise.all(
            keys.map(async ([, isEe]) => {
              const b64 = await api.issueNoticeTemplate(isEe, version)
              const labels = await api.findPlaceholdersInDocument(b64)
              return { b64, labels }
            })
          )
          if (cancelled) return
          setTemplates(Object.fromEntries(keys.map(([key], i) => [key, slots[i]])))
        } catch (e) {
          if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e))
        }
      })()
    }

    load()
    // Issue Notices stays mounted (KeepAlive) across tab switches, so without
    // this listener a variant picked in Settings after this tab was first
    // opened would never be reflected — see TEMPLATE_VARIANT_CHANGED_EVENT.
    function onVariantChanged(e: Event) {
      if ((e as CustomEvent).detail?.storageKey === noticeVariantKey) load()
    }
    window.addEventListener(TEMPLATE_VARIANT_CHANGED_EVENT, onVariantChanged)
    return () => {
      cancelled = true
      window.removeEventListener(TEMPLATE_VARIANT_CHANGED_EVENT, onVariantChanged)
    }
  }, [noticeVariantKey])

  function templateFor(item: PendingNoticeItem): TemplateSlot {
    return templates[item.seMode ? 'se' : 'ee'] ?? EMPTY_SLOT
  }

  // ── Manual upload path (fallback) ──────────────────────────────────────
  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  const [pdfEval, setPdfEval] = useState<TenderEvaluation | null>(null)
  const [pdfName, setPdfName] = useState('')
  const [pdfStatus, setPdfStatus] = useState<string | null>(null)
  const [placeOverride, setPlaceOverride] = useState<{ circle?: string; zone?: string }>({})
  const [uploadBusy, setUploadBusy] = useState<'notice' | 'pdf' | null>(null)
  const [uploadError, setUploadError] = useState<string | null>(null)
  useEffect(() => {
    setPlaceOverride({})
  }, [ee])
  const pdfInputRef = useRef<HTMLInputElement>(null)
  const noticeInputRef = useRef<HTMLInputElement>(null)

  const detected = useMemo(
    () => detectOffice(entries, pdfEval, notice, placeOverride),
    [entries, pdfEval, notice, placeOverride]
  )

  async function handlePdfFile(file: File) {
    setUploadError(null)
    setUploadBusy('pdf')
    setPdfStatus(null)
    setPlaceOverride({})
    try {
      const lines = await pdfToTextLinesOrOcr(file)
      const ev = parseTenderEvaluation(lines)
      if (!ev.nameOfWork && !ev.tenderId) {
        throw new Error("Couldn't read tender details from the PDF — is it the Commercial Evaluation / Stage Selected page?")
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
      setUploadError(e instanceof Error ? e.message : String(e))
    } finally {
      setUploadBusy(null)
    }
  }

  async function handleNoticeFile(file: File) {
    setUploadError(null)
    setUploadBusy('notice')
    try {
      const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
      const parsed = isPdf
        ? parseIntimationNoticeText(await pdfToTextLinesOrOcr(file))
        : parseIntimationNotice(await file.text())
      setNotice(parsed)
      setNoticeName(file.name)
    } catch (e) {
      setUploadError(e instanceof Error ? e.message : String(e))
    } finally {
      setUploadBusy(null)
    }
  }

  const workMatch = useMemo(() => (notice && pdfEval ? checkSameWork(notice, pdfEval) : null), [notice, pdfEval])
  const workMismatch = workMatch?.status === 'mismatch'
  const uploadReady = !!pdfEval && !workMismatch

  function addUploadedTile() {
    if (!pdfEval) return
    const ecv = notice?.ecvRupees ?? pdfEval.ecvRupees
    const contract =
      notice?.contractRupees ??
      pdfEval.contractRupees ??
      (ecv != null && pdfEval.tenderPercentage != null ? ecv * (1 - pdfEval.tenderPercentage / 100) : undefined)
    const item: PendingNoticeItem = {
      id: `notice-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      officeKey: undefined,
      seMode: !ee,
      zone: !ee ? detected.zone : undefined,
      circle: ee ? detected.circle : undefined,
      cno: ee ? detected.cno : undefined,
      corporation: corporationName,
      agencyName: notice?.agencyName ?? pdfEval.l1AgencyName,
      address: notice?.address,
      workName: pdfEval.nameOfWork,
      nitNo: notice?.nitNo ?? pdfEval.noticeNo,
      nitDate: notice?.nitDate ?? pdfEval.noticeDate,
      ecvRupees: ecv,
      contractRupees: contract,
      tenderPercentage: pdfEval.tenderPercentage,
      intimationDate: new Date().toISOString().slice(0, 10),
      manual: { ...EMPTY_ISSUE_NOTICE_MANUAL_FIELDS },
      done: false
    }
    onChange([...items, item])
    setNotice(null)
    setNoticeName('')
    setPdfEval(null)
    setPdfName('')
    setPdfStatus(null)
    setPlaceOverride({})
  }

  const activeItems = items.filter((it) => !it.deletedAt)

  // Group tiles by NIT No — a NIT can carry several items (e.g. multiple
  // agencies), each work with no NIT falls into its own "No NIT No." group.
  const groups = useMemo(() => {
    const map = new Map<string, PendingNoticeItem[]>()
    for (const it of activeItems) {
      const key = it.nitNo?.trim() || 'No NIT No.'
      const list = map.get(key)
      if (list) list.push(it)
      else map.set(key, [it])
    }
    return [...map.entries()].sort(([a], [b]) => a.localeCompare(b))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeItems])

  function saveItem(id: string, patch: Partial<PendingNoticeItem>) {
    onChange(items.map((it) => (it.id === id ? { ...it, ...patch } : it)))
  }
  function removeItem(id: string) {
    onChange(items.map((it) => (it.id === id ? { ...it, deletedAt: new Date().toISOString() } : it)))
  }
  function restoreItem(id: string) {
    onChange(items.map((it) => (it.id === id ? { ...it, done: false } : it)))
  }

  return (
    <div className="intimation-tool">
      <div className="intimation-tool-uploads">
        <div className="boq-actions">
          <button
            className="primary upload-btn"
            onClick={() => pdfInputRef.current?.click()}
            disabled={uploadBusy === 'pdf'}
          >
            <IconFolder /> {uploadBusy === 'pdf' ? 'Reading PDF…' : pdfEval ? 'Change L1 selection form' : 'Upload L1 selection form'}
          </button>
          <button
            className="primary upload-btn"
            onClick={() => noticeInputRef.current?.click()}
            disabled={uploadBusy === 'notice'}
          >
            <IconFolder /> {uploadBusy === 'notice' ? 'Reading…' : notice ? 'Change Online Intimation' : 'Upload Online Intimation (address)'}
          </button>
          {uploadReady && (
            <button className="primary" onClick={addUploadedTile}>
              <IconCheck /> Add to Issue Notices
            </button>
          )}
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
          Notices for an issued Intimation appear below automatically — use this only for a work issued before this
          existed, or from another source.
        </p>
        {pdfName && <p className="estimate-hint">Work &amp; tender details read from {pdfName}</p>}
        {noticeName && <p className="estimate-hint">Address read from {noticeName}</p>}
        {pdfStatus && (
          <div className="notice ok">
            <IconBell /> {pdfStatus}
          </div>
        )}
        {uploadError && (
          <div className="notice error">
            <IconWarn /> {uploadError}
          </div>
        )}
        {workMismatch && workMatch && (
          <div className="notice error">
            <IconWarn /> {sameWorkMismatchMessage(workMatch)}
          </div>
        )}
        {pdfEval && (
          <div className="loa-manual-fields">
            <span className="estimate-preview-title">Office (letterhead name)</span>
            <div className="loa-manual-grid">
              {ee ? (
                <label className="wo-date-field">
                  <span>Circle</span>
                  <select value={detected.circle} onChange={(e) => setPlaceOverride({ circle: e.target.value })}>
                    <option value="">—</option>
                    {entries.map((c) => (
                      <option key={c.circle} value={c.circle}>
                        {c.circle} {c.cno} ({c.zone})
                      </option>
                    ))}
                  </select>
                </label>
              ) : (
                <label className="wo-date-field">
                  <span>Zone</span>
                  <select value={detected.zone} onChange={(e) => setPlaceOverride({ zone: e.target.value })}>
                    <option value="">—</option>
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
      </div>

      {loadError && (
        <div className="notice error">
          <IconWarn /> {loadError}
        </div>
      )}

      {groups.map(([nitNo, groupItems]) => (
        <div key={nitNo} className="notice-nit-group">
          <span className="estimate-preview-title">NIT No. {nitNo}</span>
          <div className="wo-tiles">
            {groupItems.map((item) => (
              <NoticeTile
                key={item.id}
                item={item}
                templateB64={templateFor(item).b64}
                labels={templateFor(item).labels}
                onSave={(patch) => saveItem(item.id, patch)}
                onRemove={() => removeItem(item.id)}
                onRestore={() => restoreItem(item.id)}
              />
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}

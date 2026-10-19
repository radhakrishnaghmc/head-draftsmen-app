import { useEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { parseIntimationNotice, parseIntimationNoticeText, type IntimationNotice } from '@core/intimationNotice'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import { checkSameWork, sameWorkMismatchMessage } from '@core/sameWorkCheck'
import { circleFromNit } from '@core/workOrderAgreement'
import { variantForCorporation, corporationDefaultVariant } from '@core/workOrderTemplateVariants'
import {
  issueNoticePlaceholders,
  isNoticeOverdue,
  defaultLoaNo,
  EMPTY_ISSUE_NOTICE_MANUAL_FIELDS,
  type IssueNoticeManualFields,
  type IssueNoticeOffice
} from '@core/issueNotice'
import type { PendingNoticeItem, ExcelTable } from '@core/types'
import { computeWorkAmounts, tenderPercentFromRow, lakhsToRupees, formatRupees } from '@core/worksAmounts'
import { findWorksRowByName } from '@core/scheduleA'
import { agencyKey, canonicalAgencyNames } from '@core/agencyName'
import { inferTypeOfWork } from '@core/workTypeClassifier'
import { parsePct } from '@core/intimationFill'
import { stripItemNoTag, extractItemNo } from '@core/bidDocument'
import DocFieldsEditor, { groupLabels } from './DocFieldsEditor'
import { CMC_ZONE_CIRCLES, resolveFromDirectory, corporationByName } from '../zoneCircleDirectory'
import { isZoneOnlyOffice, officeScopedKey, TEMPLATE_KEYS, TEMPLATE_VARIANT_CHANGED_EVENT, type Office } from '../office'
import type { PlaceholderMatch } from '@core/createDocument'
import { pdfToTextLinesOrOcr } from '../pdfToText'
import { base64ToUint8, PAGE_WIDTH, renderDocPreview } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell, IconTrash, IconCheck, IconRestore, IconChevronLeft } from './Icons'

interface Props {
  /** chosen office — decides which bundled format the manual-upload path
   * defaults to (Zone-level SE vs. Circle-level EE), same rule every tile
   * uses for the work it belongs to. */
  office?: Office
  /** Works List (tables[0]) — a third way to start a Notice: pick the work, its row fills the fields. */
  worksTable?: ExcelTable | null
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
  /** The always-present "New Notice" tile: nothing is saved until Add to Issue Notices. */
  draft?: boolean
  onAddDraft?: () => void
  /** What the tile's footer lists: agency, type of work and estimate amount. */
  summary: NoticeSummary
}

/** The three lines a Notice tile shows in place of the bare work name. */
interface NoticeSummary {
  agency: string
  typeOfWork: string
  /** Estimate amount in rupees (the Works List's "Amount of estimate" — not the ECV), when known. */
  estimateRupees: number | null
}

/**
 * Agency, type of work and estimate amount for a notice. The type and estimate
 * come from the work's Works List row (found by name) when it has one; the type
 * otherwise comes from the same work-type classifier the Issue Tender Notice
 * document uses, and the estimate stays blank rather than borrowing the ECV.
 */
function noticeSummary(item: PendingNoticeItem, worksTable: ExcelTable | null | undefined): NoticeSummary {
  const row = item.workName && worksTable ? findWorksRowByName(worksTable, item.workName)?.row : undefined
  const listed = (row?.['Type of work'] ?? '').trim()
  const estimate = row ? lakhsToRupees(row['Amount of estimate'] ?? '') : 0
  return {
    agency: (item.agencyName ?? '').trim(),
    typeOfWork: listed || (item.workName ? (inferTypeOfWork(item.workName) ?? '') : ''),
    estimateRupees: estimate > 0 ? estimate : null
  }
}

const MANUAL_LABEL_KEYS: Record<string, keyof IssueNoticeManualFields> = {
  'LOA No': 'loaNo',
  'LOA Date': 'loaDate',
  'Notice Date': 'noticeDate',
  'agency phone number': 'agencyPhone'
}

function NoticeTile({ item, templateB64, labels, onSave, onRemove, onRestore, draft, onAddDraft, summary }: NoticeTileProps) {
  const [expanded, setExpanded] = useState(false)
  const [busy, setBusy] = useState<'download' | 'pdf' | 'print' | null>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [previewPages, setPreviewPages] = useState(1)
  const [tileFailed, setTileFailed] = useState(false)
  const tileRef = useRef<HTMLDivElement>(null)
  const expandedRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  const derived = useMemo(() => placeholdersForItem(item), [item])
  const edits = item.edits ?? {}
  const values = useMemo(() => ({ ...derived, ...(item.edits ?? {}) }), [derived, item.edits])
  const groups = useMemo(() => groupLabels(labels), [labels])
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

  // LOA No / LOA Date / Notice Date / phone live in `manual` (the overdue blink reads Notice Date); everything else is a plain override.
  function changeField(label: string, value: string) {
    const key = MANUAL_LABEL_KEYS[label]
    if (key) onSave({ manual: { ...item.manual, [key]: value } })
    else onSave({ edits: { ...edits, [label]: value } })
  }
  function revertField(label: string) {
    const next = { ...edits }
    delete next[label]
    onSave({ edits: next })
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
      await renderDocPreview(base64ToUint8(filled), container, { forSnapshot: true })
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
      <div
        className={`wo-tile ${isNoticeOverdue(item.manual.noticeDate, item.done) ? 'wo-tile-overdue' : ''}`}
        role="button"
        tabIndex={0}
        onClick={() => setExpanded(true)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault()
            setExpanded(true)
          }
        }}
        title="Notice for Non-conclusion of Agreement"
      >
        {draft && <span className="notice-draft-badge">New notice</span>}
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
        <div className="wo-tile-foot notice-tile-foot">
          {!summary.agency && !summary.typeOfWork && !item.workName ? (
            <div>{draft ? 'New Notice — click to fill in' : 'Notice for Non-conclusion of Agreement'}</div>
          ) : (
            <>
              <div>
                <span>Name of the agency:</span> {summary.agency || '—'}
              </div>
              <div>
                <span>Type of work:</span> {summary.typeOfWork || '—'}
              </div>
              <div>
                <span>Estimate amount:</span> {summary.estimateRupees != null ? formatRupees(summary.estimateRupees) : '—'}
              </div>
            </>
          )}
          {item.done && <div className="notice-tile-resolved">Resolved</div>}
        </div>
        {!draft && (
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
        )}
      </div>

      {expanded &&
        createPortal(
          <div className="wo-modal-overlay" onClick={() => setExpanded(false)}>
            <div className="wo-modal wide doc-edit-modal" onClick={(e) => e.stopPropagation()}>
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
                  <div className="ns-body">
                    <DocFieldsEditor
                      groups={groups}
                      derived={derived}
                      edits={edits}
                      onChange={changeField}
                      onRevert={revertField}
                      onResetAll={() => onSave({ edits: {} })}
                    />
                    <div className="ns-preview">
                      <div className="ns-preview-scroll">
                        <div ref={expandedRef} className="intimation-docx-preview" />
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              <div className="wo-modal-foot">
                {actionError && (
                  <div className="notice error">
                    <IconWarn /> {actionError}
                  </div>
                )}
                {actionSaved && <span className="estimate-hint">{actionSaved}</span>}
                {draft && (
                  <button
                    className="primary"
                    onClick={() => {
                      onAddDraft?.()
                      setExpanded(false)
                    }}
                    disabled={!item.workName && !item.agencyName}
                    title="Keep this Notice in your Issue Notices list"
                  >
                    <IconCheck /> Add to Issue Notices
                  </button>
                )}
                <button className="ghost wo-foot-close" onClick={() => setExpanded(false)}>
                  Close
                </button>
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
function AgreementNotConcludedWorkspace({ office, worksTable, items, onChange, view }: Props & { view: 'issue' | 'reminders' }) {
  const issueView = view === 'issue'
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
          const variant = variantForCorporation(
            localStorage.getItem(noticeVariantKey) ?? corporationDefaultVariant(office?.corporation, undefined),
            office?.corporation
          )
          const slots = await Promise.all(
            keys.map(async ([, isEe]) => {
              const b64 = await api.issueNoticeTemplate(isEe, variant)
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
  }, [noticeVariantKey, office?.corporation])

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
    const zone = !ee ? detected.zone : undefined
    const circle = ee ? detected.circle : undefined
    const cno = ee ? detected.cno : undefined
    // LOA Date = the L-1 sheet's own "Server Time" footer (the date the LOA
    // was actually issued) — same source Give Intimation seeds the Notice's
    // LOA Date from (core/issueNotice.ts's "loa date" case) — so the Notice
    // never starts with a blank Ref line the office has to look up by hand.
    const loaDate = pdfEval.serverDate || ''
    const loaNo = defaultLoaNo({ zone, circle, cno, corporation: corporationName }, loaDate)
    const item: PendingNoticeItem = {
      id: `notice-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      officeKey: undefined,
      seMode: !ee,
      zone,
      circle,
      cno,
      corporation: corporationName,
      agencyName: notice?.agencyName ?? pdfEval.l1AgencyName,
      address: notice?.address,
      workName: pdfEval.nameOfWork,
      nitNo: notice?.nitNo ?? pdfEval.noticeNo,
      nitDate: notice?.nitDate ?? pdfEval.noticeDate,
      itemNo: pdfEval.itemNo,
      ecvRupees: ecv,
      contractRupees: contract,
      tenderPercentage: pdfEval.tenderPercentage,
      intimationDate: new Date().toISOString().slice(0, 10),
      manual: { ...EMPTY_ISSUE_NOTICE_MANUAL_FIELDS, loaDate, loaNo },
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

  // ── Works List path ────────────────────────────────────────────────────
  // A third way in: pick a work from the (office-scoped) Works List and its
  // row fills the Notice; whatever the row doesn't carry stays editable in the
  // preview. Also doubles as the always-visible "New Notice" tile's draft.
  const blankDraft = (): PendingNoticeItem => ({
    id: `notice-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    officeKey: undefined,
    seMode: !ee,
    zone: !ee ? office?.zone : undefined,
    circle: ee ? office?.circle : undefined,
    cno: ee ? office?.circleNumber : undefined,
    corporation: corporationName,
    intimationDate: new Date().toISOString().slice(0, 10),
    manual: { ...EMPTY_ISSUE_NOTICE_MANUAL_FIELDS },
    done: false
  })
  const [draft, setDraft] = useState<PendingNoticeItem>(blankDraft)
  const [workQuery, setWorkQuery] = useState('')
  const [workPickerOpen, setWorkPickerOpen] = useState(false)
  const [pickedRow, setPickedRow] = useState<number | null>(null)
  const nameHeader = worksTable?.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  // Spelling variants of one contractor ("M/S AHR Constructions" / "AHR CONSTRUCTIONS")
  // share one display name, so a search for the agency finds all its works together.
  const canonicalAgency = useMemo(
    () => canonicalAgencyNames((worksTable?.rows ?? []).map((r) => r['Name of the Agency'])),
    [worksTable]
  )
  const workOptions = useMemo(() => {
    if (!worksTable || !nameHeader) return []
    const q = workQuery.trim().toLowerCase()
    const qKey = agencyKey(workQuery)
    const out: { index: number; name: string; agency: string }[] = []
    worksTable.rows.forEach((row, index) => {
      const name = (row[nameHeader] ?? '').trim()
      if (!name) return
      const raw = (row['Name of the Agency'] ?? '').trim()
      const agency = canonicalAgency.get(agencyKey(raw)) ?? raw
      const agencyHit = !!q && (agency.toLowerCase().includes(q) || raw.toLowerCase().includes(q) || (qKey !== '' && agencyKey(raw).includes(qKey)))
      if (q && !name.toLowerCase().includes(q) && !agencyHit && !(row['Wincode'] ?? '').toLowerCase().includes(q)) return
      out.push({ index, name, agency })
    })
    return out
  }, [worksTable, nameHeader, workQuery, canonicalAgency])
  // When the search text names an agency, the distinct agencies it matches: the
  // list is then every work of that agency (uncapped), headed by a summary.
  const matchedAgencies = useMemo(() => {
    const q = workQuery.trim().toLowerCase()
    const qKey = agencyKey(workQuery)
    if (!q) return []
    const hit = (a: string) => a.toLowerCase().includes(q) || (qKey !== '' && agencyKey(a).includes(qKey))
    return Array.from(new Set(workOptions.map((o) => o.agency).filter((a) => a && hit(a))))
  }, [workOptions, workQuery])

  function pickWork(index: number) {
    const row = worksTable?.rows[index]
    if (!row) return
    const amounts = computeWorkAmounts(row)
    const tenderPercentage = parsePct(tenderPercentFromRow(row))
    const ecv = amounts.ecv ?? undefined
    const rawName = row[nameHeader ?? ''] ?? ''
    const base = blankDraft()
    const zone = !ee ? office?.zone || row['Zone'] || undefined : undefined
    const circle = ee ? office?.circle || row['Circle'] || undefined : undefined
    const cno = ee ? office?.circleNumber || row['CNO'] || undefined : undefined
    const loaDate = ''
    setDraft({
      ...base,
      zone,
      circle,
      cno,
      agencyName: canonicalAgency.get(agencyKey(row['Name of the Agency'])) || row['Name of the Agency'] || undefined,
      address: row['Address of the agency'] || undefined,
      workName: stripItemNoTag(rawName),
      nitNo: row['Tender Notice No'] || undefined,
      itemNo: extractItemNo(rawName) ?? undefined,
      ecvRupees: ecv,
      contractRupees: amounts.contractAmount ?? undefined,
      tenderPercentage,
      manual: {
        ...EMPTY_ISSUE_NOTICE_MANUAL_FIELDS,
        loaDate,
        loaNo: defaultLoaNo({ zone, circle, cno, corporation: corporationName }, loaDate)
      }
    })
    setPickedRow(index)
    setWorkQuery('')
    setWorkPickerOpen(false)
  }

  function addDraftTile() {
    onChange([...items, draft])
    setDraft(blankDraft())
    setPickedRow(null)
  }

  const activeItems = items.filter((it) => !it.deletedAt)
  const activeCount = activeItems.filter((it) => !it.done).length
  const resolvedCount = activeItems.length - activeCount
  const overdueCount = activeItems.filter((it) => isNoticeOverdue(it.manual.noticeDate, it.done)).length

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
    <div className="intimation-tool ntc-workspace">
      {issueView && (
      <aside className="ntc-side">
      <section className="ntc-compose" aria-label="Create a notice">
        <header className="ntc-compose-head">
          <h3>Create a notice</h3>
          <p>Start from a work in the Works List, or read the details from the tender documents.</p>
        </header>
        <div className="ntc-sources">
        {worksTable && nameHeader && (
          <div className="ntc-source notice-work-picker">
            <span className="ntc-source-title">
              <span className="ntc-step">1</span> Issue from the Works List
            </span>
            <div className={`notice-work-combo${workPickerOpen ? ' open' : ''}`}>
              <input
                type="text"
                value={workQuery}
                placeholder={pickedRow != null ? `${draft.workName ?? ''}` : 'Search by work name, agency or Win Code…'}
                onFocus={() => setWorkPickerOpen(true)}
                onChange={(e) => {
                  setWorkQuery(e.target.value)
                  setWorkPickerOpen(true)
                }}
                onBlur={() => setTimeout(() => setWorkPickerOpen(false), 150)}
              />
              {workPickerOpen && (
                <div className="notice-work-list" role="listbox">
                  <div className="notice-work-source">
                    <span>Works List</span>
                    <span>
                      {workOptions.length} work{workOptions.length === 1 ? '' : 's'}
                      {!matchedAgencies.length && workOptions.length > 80 ? ' · showing first 80' : ''}
                    </span>
                  </div>
                  {matchedAgencies.length > 0 && (
                    <div className="notice-work-agency-head">
                      {matchedAgencies.length === 1 ? matchedAgencies[0] : `${matchedAgencies.length} agencies`} —{' '}
                      {workOptions.length} work{workOptions.length === 1 ? '' : 's'}
                    </div>
                  )}
                  {workOptions.length === 0 ? (
                    <div className="notice-work-empty">No matching work in this office's Works List.</div>
                  ) : (
                    (matchedAgencies.length > 0 ? workOptions : workOptions.slice(0, 80)).map((o) => (
                      <button
                        key={o.index}
                        type="button"
                        role="option"
                        className="notice-work-option"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => pickWork(o.index)}
                      >
                        <span className="notice-work-name">{o.name}</span>
                        {o.agency && <span className="notice-work-agency">{o.agency}</span>}
                      </button>
                    ))
                  )}
                </div>
              )}
            </div>
            <p className="estimate-hint">
              Fills the New Notice tile below from that work's row — then open the tile to edit anything else.
            </p>
          </div>
        )}
        <div className="ntc-source">
          <span className="ntc-source-title">
            <span className="ntc-step">{worksTable && nameHeader ? 2 : 1}</span> Upload L1 sheet &amp; Intimation
          </span>
        <div className="boq-actions boq-actions--start ntc-actions">
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
          <div className="loa-manual-fields ntc-office">
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
        </div>
      </section>
      </aside>
      )}

      <div className="ntc-board">
      {!issueView && (
      <div className="ntc-summary" aria-label="Reminders summary">
        <div className="ntc-chip">
          <span className="ntc-chip-label">Active</span>
          <span className="ntc-chip-value">{activeCount}</span>
        </div>
        <div className={`ntc-chip${overdueCount > 0 ? ' ntc-chip-warn' : ''}`}>
          <span className="ntc-chip-label">Overdue</span>
          <span className="ntc-chip-value">{overdueCount}</span>
        </div>
        <div className="ntc-chip">
          <span className="ntc-chip-label">Resolved</span>
          <span className="ntc-chip-value">{resolvedCount}</span>
        </div>
      </div>
      )}


      {loadError && (
        <div className="notice error">
          <IconWarn /> {loadError}
        </div>
      )}

      {issueView && (
      <div className="notice-nit-group">
        <div className="ntc-group-head">
          <span className="ntc-group-title">New Notice</span>
        </div>
        <div className="wo-tiles">
          <NoticeTile
            key={draft.id}
            item={draft}
            draft
            templateB64={templateFor(draft).b64}
            labels={templateFor(draft).labels}
            onSave={(patch) => setDraft((d) => ({ ...d, ...patch }))}
            onRemove={() => {}}
            onRestore={() => {}}
            onAddDraft={addDraftTile}
            summary={noticeSummary(draft, worksTable)}
          />
        </div>
        <p className="estimate-hint">
          Notices you issue here are tracked as reminders under “Agreement not Concluded — Reminders”.
        </p>
      </div>
      )}

      {!issueView && groups.length === 0 && (
        <div className="empty">
          <IconBell />
          <p>No reminders yet. A notice reminder is added here automatically when an Intimation is issued.</p>
        </div>
      )}

      {!issueView && groups.map(([nitNo, groupItems]) => (
        <div key={nitNo} className="notice-nit-group">
          <div className="ntc-group-head">
            <span className="ntc-group-title">NIT No. {nitNo}</span>
            <span className="ntc-group-count">{groupItems.length}</span>
          </div>
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
                summary={noticeSummary(item, worksTable)}
              />
            ))}
          </div>
        </div>
      ))}
      </div>
    </div>
  )
}

/**
 * Issue Notices landing page: one tile per kind of notice. Only "Agreement not
 * Concluded" exists today — add further kinds to NOTICE_KINDS and render their
 * workspace below. The opened workspace stays mounted (hidden) on Back so an
 * in-progress draft / uploaded documents aren't lost.
 */
const NOTICE_KINDS = [
  {
    key: 'agreementNotConcluded' as const,
    name: 'Agreement not Concluded',
    meta: 'Issue the Notice for non-conclusion of the Agreement — from the Works List, or by uploading the L1 sheet and Intimation',
    view: 'issue' as const
  },
  {
    key: 'agreementNotConcludedReminders' as const,
    name: 'Agreement not Concluded — Reminders',
    meta: 'Notices prepared automatically when an Intimation is issued, with their due dates',
    view: 'reminders' as const
  }
]
type NoticeKind = (typeof NOTICE_KINDS)[number]['key']

export default function IssueNoticesTab(props: Props & { onWorkspaceChange?: (open: boolean) => void }) {
  const { onWorkspaceChange, ...workspaceProps } = props
  const [kind, setKind] = useState<NoticeKind | null>(null)
  const [opened, setOpened] = useState<Set<NoticeKind>>(new Set())
  // Bumping this remounts the issue workspace, clearing uploads and typed fields.
  const [resetKey, setResetKey] = useState(0)
  useEffect(() => {
    onWorkspaceChange?.(kind !== null)
  }, [kind, onWorkspaceChange])

  const active = props.items.filter((it) => !it.deletedAt && !it.done)
  const overdue = active.filter((it) => isNoticeOverdue(it.manual.noticeDate, it.done)).length

  function open(k: NoticeKind) {
    setOpened((prev) => new Set(prev).add(k))
    setKind(k)
  }

  return (
    <>
      {kind === null && (
        <div className="card">
          <div className="doc-tile-grid tools-grid notice-kind-grid">
            {NOTICE_KINDS.map((k) => (
              <button key={k.key} className="doc-tile-card tone-green tool-card" onClick={() => open(k.key)}>
                <span className="tool-card-ic">
                  <IconBell />
                </span>
                <span className="doc-tile-card-name">{k.name}</span>
                <span className="doc-tile-card-meta">
                  {k.meta}
                  {k.view === 'reminders' && ` · ${active.length} active${overdue ? `, ${overdue} overdue` : ''}`}
                </span>
                <span className="tool-card-cta">
                  <IconFolder /> Open workspace
                </span>
              </button>
            ))}
          </div>
        </div>
      )}
      {NOTICE_KINDS.filter((k) => opened.has(k.key)).map((k) => (
        <div key={k.key} hidden={kind !== k.key}>
          <div className="pdf-ws-topbar">
            <div className="pdf-ws-title">
              <IconBell /> <span className="pdf-ws-title-label">Notice:</span> {k.name}
            </div>
            {k.view === 'issue' && (
              <button
                className="ghost"
                onClick={() => setResetKey((n) => n + 1)}
                title="Clear the uploaded documents and filled fields to start a new work"
              >
                <IconRestore /> Clear
              </button>
            )}
            <button className="ghost pdf-ws-back" onClick={() => setKind(null)}>
              <IconChevronLeft /> Back
            </button>
          </div>
          <div className="card pdf-workspace">
            <AgreementNotConcludedWorkspace
              key={k.view === 'issue' ? resetKey : 'reminders'}
              {...workspaceProps}
              view={k.view}
            />
          </div>
        </div>
      ))}
    </>
  )
}

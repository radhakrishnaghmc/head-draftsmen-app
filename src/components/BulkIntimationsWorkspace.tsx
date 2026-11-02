import { useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { pdfToTextLinesOrOcr } from '../pdfToText'
import { parseTenderEvaluation, type TenderEvaluation } from '@core/tenderEvaluationPdf'
import {
  parseIntimationNotice,
  parseIntimationNoticeText,
  parseSupplierDetailsText,
  type IntimationNotice
} from '@core/intimationNotice'
import { detectTenderId } from '@core/tenderAgents/tenderId'
import type { ParsedDoc } from '@core/massNotices'
import { pairForIntimations } from '@core/bulkIntimations'
import { deriveFields, qcIntimationCompletionPlaceholders, qcPartyToBlock, aeDyeeForwardingPlaceholders } from '@core/workOrderAgreement'
import { checkSameWork } from '@core/sameWorkCheck'
import { resolveIntimationValue } from '@core/intimationFill'
import { matchWorksRowCached, buildWorksRowEmbeddingCache } from '../estimateDownloads'
import { inferTypeOfWork } from '@core/workTypeClassifier'
import { lakhsToRupees, formatRupees } from '@core/worksAmounts'
import type { ExcelTable, QcOfficeParties, StoredLetter } from '@core/types'
import type { PlaceholderMatch } from '@core/createDocument'
import {
  variantForCorporation,
  corporationDefaultVariant,
  INTIMATION_TEMPLATE_VARIANTS,
  QC_INTIMATION_COMPLETION_TEMPLATE_VARIANTS
} from '@core/workOrderTemplateVariants'
import { CMC_ZONE_CIRCLES, corporationByName } from '../zoneCircleDirectory'
import { isZoneOnlyOffice, officeScopedKey, TEMPLATE_KEYS, type Office } from '../office'
import { closeOnBackdropMouseDown } from '../overlayClose'
import { detectOffice } from './IssueNoticesTab'
import { base64ToUint8, renderDocPreview } from './docPage'
import {
  IconFolder,
  IconWarn,
  IconCheck,
  IconDoc,
  IconChevronLeft,
  IconRestore,
  IconDownload,
  IconTrash
} from './Icons'

interface Props {
  office: Office
  /** The office's Works List — gives each tile its type of work / estimate amount and fills the letter's estimate. */
  worksTable?: ExcelTable | null
  /**
   * 'qc' makes the five-letter QC Intimation/Completion set and 'forwarding' the AE/Dy.EE Forwarding Letters instead of the
   * Intimation. Those two are handed to `onAddLetters` — they land in the Forwarding Slips & QC Intimations tab.
   */
  kind?: 'intimation' | 'qc' | 'forwarding'
  /** Receives every letter a 'qc' / 'forwarding' run made. */
  onAddLetters?: (letters: StoredLetter[]) => void
  /** The office's 3rd/4th-party QC agencies (Works List agency panel) — addressed on the QC letters. */
  qcParties?: QcOfficeParties
  /** Leaves the workspace — the parent unmounts it, which resets it to the Upload folder state. */
  onBack: () => void
}

/** One ready Intimation: the filled letter plus what its tile shows. */
interface ReadyIntimation {
  id: string
  docx: string
  nitNo: string
  nitDate: string
  agency: string
  typeOfWork: string
  estimateRupees: number | null
  fileName: string
}

interface Rejected {
  key: string
  workName: string
  nitNo: string
  agency: string
  note: string
}

type Phase = 'idle' | 'opening' | 'reading' | 'generating' | 'done'

const pause = (ms: number) => new Promise<void>((r) => setTimeout(r, ms))

/** Thumbnail of one filled letter — the same live preview every other document tile uses. */
function TilePreview({ docx, full = false }: { docx: string; full?: boolean }) {
  const ref = useRef<HTMLDivElement>(null)
  const [failed, setFailed] = useState(false)
  useEffect(() => {
    const el = ref.current
    if (!el) return
    let cancelled = false
    void renderDocPreview(base64ToUint8(docx), el)
      .then(() => !cancelled && setFailed(false))
      .catch(() => !cancelled && setFailed(true))
    return () => {
      cancelled = true
    }
  }, [docx])
  return failed ? (
    <div className="settings-tile-error">
      <IconWarn /> Couldn't load
    </div>
  ) : full ? (
    // Full-size page for the preview modal — the tile's own container is a scaled, absolutely positioned thumbnail.
    <div className="ns-preview">
      <div className="ns-preview-scroll">
        <div ref={ref} className="intimation-docx-preview" />
      </div>
    </div>
  ) : (
    <div ref={ref} className="wo-tile-doc" />
  )
}

/**
 * Settings → Bulk Intimations workspace. Pick a folder of tender documents; each L-1 selection sheet is paired with
 * its LOA and an Intimation is prepared ONLY when the LOA's NIT number, agency name and contract amount all agree
 * with the L-1 sheet (core/bulkIntimations.ts) — a pair that fails any of the three is listed with the reason
 * instead. Ready letters show as tiles like Issue Notices; Download all saves them (Word or PDF) into one folder and,
 * once saved, clears them from the workspace.
 */
export default function BulkIntimationsWorkspace({ office, worksTable, kind = 'intimation', qcParties, onAddLetters, onBack }: Props) {
  const isQc = kind === 'qc'
  const isLetters = kind !== 'intimation'
  const title = isQc ? 'Bulk QC Intimations' : kind === 'forwarding' ? 'Bulk Forwarding Letters' : 'Bulk Intimations'
  const docName = isQc ? 'QC Intimation/Completion' : kind === 'forwarding' ? 'AE/Dy.EE Forwarding' : 'Intimation'
  const destination = isQc ? 'QC Intimation/Completion Letters' : 'AE/Dy.EE Forwarding Letters'
  const [phase, setPhase] = useState<Phase>('idle')
  const [readProgress, setReadProgress] = useState({ done: 0, total: 0 })
  const [scan, setScan] = useState<{ phase: 'scanning' | 'reading'; done: number; total: number } | null>(null)
  const openingRef = useRef(false)
  const [genProgress, setGenProgress] = useState({ done: 0, total: 0 })
  const [error, setError] = useState<string | null>(null)
  const [tiles, setTiles] = useState<ReadyIntimation[]>([])
  // Letters handed to the Forwarding Slips & QC Intimations tab (they aren't kept here).
  const [madeCount, setMadeCount] = useState(0)
  const [rejected, setRejected] = useState<Rejected[]>([])
  const [ignored, setIgnored] = useState(0)
  const [format, setFormat] = useState<'docx' | 'pdf'>('docx')
  const [saving, setSaving] = useState<{ done: number; total: number } | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const [preview, setPreview] = useState<ReadyIntimation | null>(null)

  const cancelled = useRef(false)
  const runId = useRef(0)
  useEffect(
    () =>
      api.onTenderScanProgress((p) => {
        if (openingRef.current && p.phase !== 'parsing') setScan({ phase: p.phase, done: p.done, total: p.total })
      }),
    []
  )
  useEffect(() => {
    cancelled.current = false
    return () => {
      cancelled.current = true
    }
  }, [])

  const seOffice = isZoneOnlyOffice(office)
  const corp = corporationByName(office.corporation)
  const entries = corp?.entries?.length ? corp.entries : CMC_ZONE_CIRCLES
  const busy = phase === 'opening' || phase === 'reading' || phase === 'generating' || !!saving

  function clear() {
    runId.current++
    openingRef.current = false
    setPhase('idle')
    setError(null)
    setTiles([])
    setMadeCount(0)
    setRejected([])
    setIgnored(0)
    setScan(null)
    setSaved(null)
    setReadProgress({ done: 0, total: 0 })
    setGenProgress({ done: 0, total: 0 })
  }

  async function run() {
    const myRun = ++runId.current
    const stale = () => cancelled.current || runId.current !== myRun
    setPhase('opening')
    setScan(null)
    openingRef.current = true
    setError(null)
    setTiles([])
    setMadeCount(0)
    setRejected([])
    setIgnored(0)
    setSaved(null)
    try {
      // The office's own Intimation format, as Give Intimation / the Intimation tool pick it.
      const variant = variantForCorporation(
        localStorage.getItem(officeScopedKey(TEMPLATE_KEYS.intimation, office)) ??
          corporationDefaultVariant(office.corporation, undefined),
        office.corporation,
        INTIMATION_TEMPLATE_VARIANTS
      )
      const qcVariant = variantForCorporation(
        localStorage.getItem(officeScopedKey(TEMPLATE_KEYS.qcIntimationCompletion, office)) ??
          corporationDefaultVariant(office.corporation, undefined),
        office.corporation,
        QC_INTIMATION_COMPLETION_TEMPLATE_VARIANTS
      )
      const fwdVariant = corporationDefaultVariant(office.corporation, undefined)
      const letterVariant = (isQc ? qcVariant : fwdVariant) ?? 'cmc'
      const templateB64 = isQc
        ? await api.qccIntimationTemplate(qcVariant)
        : kind === 'forwarding'
          ? await api.aeDyeeForwardingTemplate(fwdVariant)
          : await api.intimationTemplate(variant)
      const labels = await api.findPlaceholdersInDocument(templateB64)

      const picked = await api.pickTenderDocuments().finally(() => {
        openingRef.current = false
      })
      setScan(null)
      if (stale()) return
      if (picked.length === 0) {
        setPhase('idle')
        return
      }

      // 1) Classify every file as an L-1 sheet or an LOA (Online Intimation).
      setPhase('reading')
      const l1s: ParsedDoc<TenderEvaluation>[] = []
      const loas: ParsedDoc<IntimationNotice>[] = []
      let unread = 0
      for (let i = 0; i < picked.length; i++) {
        if (stale()) return
        const f = new File([new Uint8Array(picked[i].bytes)], picked[i].name)
        setReadProgress({ done: i + 1, total: picked.length })
        try {
          if (/\.pdf$/i.test(f.name)) {
            const lines = await pdfToTextLinesOrOcr(f)
            const ev = parseTenderEvaluation(lines)
            if ((ev.nameOfWork || ev.tenderId) && ev.l1AgencyName) {
              l1s.push({ file: f.name, doc: ev })
              continue
            }
            const n = parseIntimationNoticeText(lines)
            // A Supplier Details page is not an LOA — it has no NIT / contract amount to verify against.
            if (!n.nitNo && detectTenderId(lines) && parseSupplierDetailsText(lines).agencyName) {
              unread++
              continue
            }
            if (n.agencyName && n.nitNo) loas.push({ file: f.name, doc: n })
            else unread++
          } else {
            const n = parseIntimationNotice(await f.text())
            if (n.agencyName && n.nitNo) loas.push({ file: f.name, doc: n })
            else unread++
          }
        } catch {
          unread++
        }
      }
      setIgnored(unread)
      if (l1s.length === 0) {
        setError("Couldn't find any L-1 selection sheet (Commercial Evaluation / Stage Selected PDF) in that folder.")
        setPhase('idle')
        return
      }

      // 2) Only this office's tenders (by the work's own circle), then the NIT / agency / contract-amount match.
      const mine = (office.circle ?? '').trim().toLowerCase()
      const own = l1s.filter((x) => (detectOffice(entries, x.doc, null, {}).circle ?? '').trim().toLowerCase() === mine)
      // Other circles' L-1 sheets and their LOAs (same NIT) aren't ours to process — count them as ignored, not rejected.
      const foreign = l1s.filter((y) => !own.includes(y))
      const foreignLoas = new Set(
        loas.filter((l) => foreign.some((x) => checkSameWork({ nitNo: l.doc.nitNo }, x.doc).status === 'match'))
      )
      if (foreign.length || foreignLoas.size) setIgnored(unread + foreign.length + foreignLoas.size)
      const ownLoas = loas.filter((l) => !foreignLoas.has(l) || own.some((x) => checkSameWork({ nitNo: l.doc.nitNo }, x.doc).status === 'match'))
      const { pairs, rejected: bad } = pairForIntimations(own, ownLoas)
      for (const b of bad)
        setRejected((r) => [
          ...r,
          {
            key: `bad|${b.l1File}`,
            workName: b.l1.nameOfWork ?? '',
            nitNo: b.l1.noticeNo ?? '',
            agency: b.l1.l1AgencyName ?? '',
            note: b.loaFile ? `${b.reason} (LOA ${b.loaFile})` : b.reason
          }
        ])

      // 3) One letter per verified pair.
      setPhase('generating')
      setGenProgress({ done: 0, total: pairs.length })
      const intimationOffice = { ...office, corporationFullName: corporationByName(office.corporation)?.fullName }
      // Exact work-name match first, embedding fallback second — the L-1 PDF's wording often differs slightly from the Works List's.
      const rowCache = worksTable ? await buildWorksRowEmbeddingCache(worksTable) : null
      const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
      const seen = new Set<string>()
      const made: StoredLetter[] = []
      for (let idx = 0; idx < pairs.length; idx++) {
        if (stale()) return
        const { l1, loa } = pairs[idx]
        setGenProgress({ done: idx + 1, total: pairs.length })
        const dupKey = `${l1.tenderId ?? ''}|${l1.noticeNo ?? ''}|${l1.nameOfWork ?? ''}`
        if (seen.has(dupKey)) continue
        seen.add(dupKey)
        try {
          const dbRow =
            l1.nameOfWork && worksTable ? (await matchWorksRowCached(l1.nameOfWork, worksTable, rowCache))?.row : undefined
          const row: Record<string, string> = {
            ...(dbRow ?? {}),
            Circle: office.circle ?? '',
            CNO: office.circleNumber ?? '',
            Zone: office.zone ?? '',
            Corporation: office.corporation ?? '',
            'Corp Full': corporationByName(office.corporation)?.fullName ?? '',
            'Name of the work': l1.nameOfWork ?? '',
            'Tender Notice No': l1.noticeNo || loa.nitNo || ''
          }
          let values: Record<string, string> = {}
          if (isLetters) {
            // The same work fields the Agreement workspace builds, then this letter's own placeholder map.
            const fields = {
              ...deriveFields(loa, l1, row),
              circle: office.circle ?? '',
              cno: office.circleNumber ?? '',
              zone: office.zone ?? '',
              corporation: office.corporation ?? '',
              corporationFullName: corporationByName(office.corporation)?.fullName ?? ''
            }
            values = isQc
              ? qcIntimationCompletionPlaceholders({
                  ...fields,
                  qcDepartmentTo: qcPartyToBlock(qcParties?.department),
                  qcThirdTo: qcPartyToBlock(qcParties?.third),
                  qcFourthTo: qcPartyToBlock(qcParties?.fourth)
                })
              : aeDyeeForwardingPlaceholders(fields)
          } else {
            for (const label of labels) values[label] = resolveIntimationValue(label, loa, l1, row, intimationOffice)
          }
          const estimate = dbRow ? lakhsToRupees(dbRow['Amount of estimate'] ?? '') : 0
          const agency = (loa.agencyName ?? l1.l1AgencyName ?? '').trim()
          const nit = (loa.nitNo ?? l1.noticeNo ?? '').trim()
          const nitForName = nit.replace(/\s*\bdated?\b.*$/i, '').replace(/[/\\:*?"<>|]/g, '-')
          const typeOfWork =
            (dbRow?.['Type of work'] ?? '').trim() || (l1.nameOfWork ? (inferTypeOfWork(l1.nameOfWork) ?? '') : '')
          const fileName = `${isQc ? 'QC Inti-Comp' : kind === 'forwarding' ? 'AE-DyEE Forwarding' : 'Intimation'}${agency ? ` - ${agency}` : ''}${nitForName ? ` - ${nitForName}` : ''}`
          if (isLetters) {
            // Kept as values, not a docx — the Forwarding Slips & QC Intimations tab re-fills the template on demand.
            made.push({
              id: `sl-${Date.now()}-${idx}`,
              kind: isQc ? 'qc' : 'forwarding',
              nitNo: nit,
              nitDate: loa.nitDate ?? l1.noticeDate ?? '',
              agency,
              workName: l1.nameOfWork ?? '',
              typeOfWork,
              estimateRupees: estimate > 0 ? estimate : null,
              variant: letterVariant,
              values,
              fileName,
              createdDate: new Date().toISOString().slice(0, 10)
            })
            setMadeCount(made.length)
            continue
          }
          const docx = await api.fillPlaceholdersInDocument(templateB64, resolved, values)
          setTiles((t) => [
            ...t,
            {
              id: `bi-${Date.now()}-${idx}`,
              docx,
              nitNo: nit,
              nitDate: loa.nitDate ?? l1.noticeDate ?? '',
              agency,
              typeOfWork,
              estimateRupees: estimate > 0 ? estimate : null,
              fileName
            }
          ])
        } catch (e) {
          setRejected((r) => [
            ...r,
            {
              key: `fail|${idx}`,
              workName: l1.nameOfWork ?? '',
              nitNo: l1.noticeNo ?? '',
              agency: l1.l1AgencyName ?? '',
              note: `Couldn't prepare the letter: ${e instanceof Error ? e.message : String(e)}`
            }
          ])
        }
        await pause(120) // let each tile paint before the next one
      }
      if (isLetters && made.length > 0) {
        onAddLetters?.(made)
        setSaved(`${made.length} letter${made.length === 1 ? '' : 's'} moved to Forwarding Slips & QC Intimations → ${destination}.`)
      }
      setPhase('done')
    } catch (e) {
      if (stale()) return
      setError(e instanceof Error ? e.message : String(e))
      setPhase('idle')
    }
  }

  // Saves every ready Intimation into one chosen folder, then clears the ones that were written.
  async function downloadAll() {
    if (tiles.length === 0) return
    setSaved(null)
    setError(null)
    const outTiles = tiles
    setSaving({ done: 0, total: tiles.length })
    const unsubscribe = api.onAgreementBundleProgress(({ done, total }) => setSaving({ done, total }))
    try {
      // Colons aren't allowed in Windows folder names, so the main process turns "Bulk Intimations: <NIT>" into "Bulk Intimations - <NIT>".
      const nits = [...new Set(outTiles
          .map((t) =>
            t.nitNo
              .replace(/[,;]?\s*tender\s*id\b.*$/i, '')
              .replace(/\s*\bdated?\b.*$/i, '')
              .replace(/\s*\/\s*/g, '/')
              .trim()
          )
          .filter(Boolean))]
      const folderName = nits.length ? `Bulk Intimations: ${nits.join(', ')}` : 'Bulk Intimations'
      const res = await api.exportAgreementBundle(outTiles.map((t) => ({ name: t.fileName, format, docxBase64: t.docx })), undefined, folderName)
      if (!res) return // folder dialog cancelled — nothing saved, nothing removed
      const failedNames = new Set(res.failed.map((f) => f.replace(/\s*\((DOCX|PDF)\)$/, '')))
      const remaining = tiles.filter((t) => failedNames.has(t.fileName))
      const savedCount = tiles.length - remaining.length
      setTiles(remaining)
      setSaved(
        remaining.length === 0
          ? `Downloaded ${savedCount} intimation${savedCount === 1 ? '' : 's'} — all removed from this workspace.`
          : `Downloaded ${savedCount}; ${remaining.length} failed and stay here to retry: ${res.failed.join(', ')}`
      )
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      unsubscribe()
      setSaving(null)
    }
  }

  const started = phase !== 'idle'
  const progress =
    phase === 'reading'
      ? {
          label: `Reading ${readProgress.done}/${readProgress.total}`,
          pct: readProgress.total ? (readProgress.done / readProgress.total) * 100 : 0
        }
      : phase === 'generating'
        ? {
            label: `Intimation ${genProgress.done}/${genProgress.total}`,
            pct: genProgress.total ? (genProgress.done / genProgress.total) * 100 : 0
          }
        : saving
          ? {
              label: `Saving ${saving.done}/${saving.total}`,
              pct: saving.total ? (saving.done / saving.total) * 100 : 0
            }
          : phase === 'opening'
            ? scan?.phase === 'reading'
              ? {
                  label: `Loading files ${scan.done}/${scan.total}`,
                  pct: scan.total ? (scan.done / scan.total) * 100 : 0
                }
              : { label: scan ? `Scanning folder… ${scan.done} found` : 'Waiting for folder…', pct: 0 }
            : null

  return (
    <>
      <div className="pdf-ws-topbar">
        <div className="pdf-ws-title">
          <IconDoc /> {title}
        </div>
        {started && (
          <button
            className="ghost"
            onClick={clear}
            disabled={!!saving}
            title="Clear this run and start again from Upload folder"
          >
            <IconRestore /> Clear
          </button>
        )}
        <button className="ghost pdf-ws-back" onClick={onBack}>
          <IconChevronLeft /> Back
        </button>
      </div>

      <div className="card pdf-workspace mass-notices-body">
        {seOffice && (
          <div className="notice warn">
            <IconWarn /> {title} makes the circle-level {docName} letters. A Zone office issues Letters of
            Acceptance from Give Intimation, which needs hand-entered details per work.
          </div>
        )}

        {!started && !seOffice && (
          <div className="mass-notices-upload">
            <p className="sub">
              Pick a folder of L-1 sheets and LOAs. {isQc ? 'The QC Intimation/Completion set' : kind === 'forwarding' ? 'An AE/Dy.EE Forwarding Letter' : 'An Intimation'} is prepared only when the LOA’s NIT number, agency
              name and contract amount match the L-1 sheet.
            </p>
            <button onClick={() => void run()}>
              <IconFolder /> Upload folder
            </button>
          </div>
        )}

        {error && (
          <div className="notice warn">
            <IconWarn /> {error}
            {phase === 'idle' && !seOffice && (
              <div>
                <button className="ghost" onClick={() => void run()} disabled={busy}>
                  <IconFolder /> Try another folder
                </button>
              </div>
            )}
          </div>
        )}

        {progress && (
          <div
            className="mass-notices-progress"
            role="progressbar"
            aria-valuenow={Math.round(progress.pct)}
            aria-valuemin={0}
            aria-valuemax={100}
          >
            <div className="mass-notices-progress-label">
              {progress.label} {progress.pct > 0 && `(${Math.round(progress.pct)}%)`}
            </div>
            <div className="mass-notices-progress-track">
              <div className="mass-notices-progress-fill" style={{ width: `${progress.pct}%` }} />
            </div>
          </div>
        )}

        {started && !error && (
          <div className="ntc-summary" aria-label="Intimations summary">
            <div className="ntc-chip">
              <span className="ntc-chip-label">Ready</span>
              <span className="ntc-chip-value">{isLetters ? madeCount : tiles.length}</span>
            </div>
            <div className={`ntc-chip${rejected.length > 0 ? ' ntc-chip-warn' : ''}`}>
              <span className="ntc-chip-label">Not made</span>
              <span className="ntc-chip-value">{rejected.length}</span>
            </div>
            {ignored > 0 && (
              <div className="ntc-chip">
                <span className="ntc-chip-label">Ignored</span>
                <span className="ntc-chip-value">{ignored}</span>
              </div>
            )}
          </div>
        )}

        {tiles.length > 0 && (
          <div className="download-all-group">
            <div className="format-toggle" role="group" aria-label="Download format">
              <button
                type="button"
                className={format === 'docx' ? 'on' : ''}
                disabled={busy}
                onClick={() => setFormat('docx')}
              >
                Word
              </button>
              <button
                type="button"
                className={format === 'pdf' ? 'on' : ''}
                disabled={busy}
                onClick={() => setFormat('pdf')}
              >
                PDF
              </button>
            </div>
            <button className="primary" onClick={() => void downloadAll()} disabled={busy}>
              <IconDownload /> {saving ? 'Saving…' : `Download all (${tiles.length})`}
            </button>
          </div>
        )}

        {saved && (
          <div className="notice">
            <IconCheck /> {saved}
          </div>
        )}
        {phase === 'done' && tiles.length === 0 && madeCount === 0 && !saved && (
          <div className="notice warn">
            <IconWarn /> No {docName} could be prepared from this folder — see “Not made” below for why.
          </div>
        )}

        {tiles.length > 0 && (
          <div className="notice-nit-group">
            <div className="wo-tiles">
              {tiles.map((t) => (
                <div
                  key={t.id}
                  className="wo-tile"
                  role="button"
                  tabIndex={0}
                  onClick={() => setPreview(t)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault()
                      setPreview(t)
                    }
                  }}
                  title={docName}
                >
                  <div className="wo-tile-preview">
                    <TilePreview docx={t.docx} />
                    <span className="wo-tile-open">Click to preview</span>
                  </div>
                  <div className="wo-tile-foot notice-tile-foot">
                    <div>
                      <span>NIT No:</span> {t.nitNo || '—'}
                    </div>
                    <div>
                      <span>NIT date:</span> {t.nitDate || '—'}
                    </div>
                    <div>
                      <span>Name of the agency:</span> {t.agency || '—'}
                    </div>
                    <div>
                      <span>Type of work:</span> {t.typeOfWork || '—'}
                    </div>
                    <div>
                      <span>Estimate amount:</span> {t.estimateRupees != null ? formatRupees(t.estimateRupees) : '—'}
                    </div>
                  </div>
                  <div className="wo-tile-actions" onClick={(e) => e.stopPropagation()}>
                    <button
                      className="ghost small"
                      onClick={() => setTiles((x) => x.filter((y) => y.id !== t.id))}
                      disabled={busy}
                      title="Leave this one out"
                    >
                      <IconTrash />
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {rejected.length > 0 && (
          <details className="mass-notices-skipped" open>
            <summary>Not made ({rejected.length})</summary>
            <table className="mass-notices-table">
              <tbody>
                {rejected.map((o) => (
                  <tr key={o.key}>
                    <td>
                      <div>{o.workName || '—'}</div>
                      <div className="estimate-hint">
                        {o.nitNo ? `NIT ${o.nitNo} · ` : ''}
                        {o.agency || 'Unknown agency'} · {o.note}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </details>
        )}
      </div>

      {preview &&
        createPortal(
          <div className="wo-modal-overlay" onMouseDown={closeOnBackdropMouseDown(() => setPreview(null))}>
            <div className="wo-modal wide" onClick={(e) => e.stopPropagation()}>
              <div className="wo-modal-head">
                <span className="wo-modal-title">{docName} — {preview.agency || preview.nitNo}</span>
                <button className="wo-modal-close" onClick={() => setPreview(null)} title="Close" aria-label="Close">
                  ×
                </button>
              </div>
              <div className="wo-modal-main">
                <div className="wo-modal-body">
                  <div className="ns-body">
                    <TilePreview docx={preview.docx} full />
                  </div>
                </div>
              </div>
              <div className="wo-modal-foot">
                <button className="ghost wo-foot-close" onClick={() => setPreview(null)}>
                  Close
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
    </>
  )
}

import { useEffect, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { api } from '../ipc'
import { IconDoc, IconEye, IconDownload, IconTrash, IconCheck, IconWarn, IconChevronRight } from './Icons'
import { base64ToUint8, renderDocPreview } from './docPage'
import type { BidDocumentBatch, BidDocumentWork } from '@core/types'
import type { BidDocumentInput, BidDocumentBatchProgress } from '../../electron/ipc-contract'
import { closeOnBackdropMouseDown } from '../overlayClose'
import { resolveTypeLabel } from '@core/tenderNotice'
import { inferWorkLocation } from '@core/workLocationExtractor'

interface Props {
  batches: BidDocumentBatch[]
  onRemove: (id: string) => void
  /** Edits one work's SE-only manual fields (Item No, TS No/Date, Admin Sanction) in place. */
  onUpdateWork: (batchId: string, serial: number, patch: Partial<BidDocumentWork>) => void
}

/** A Zone-level (SE) office has no Circle of its own — same "seMode" test used throughout the app — so the SE Bid Document template replaces the EE one for this work, rather than the two coexisting. */
function isSeWork(work: BidDocumentWork): boolean {
  return !!work.zone?.trim() && !work.circle?.trim()
}

function buildInput(batch: BidDocumentBatch, work: BidDocumentWork): BidDocumentInput {
  return {
    nitNo: batch.nitNo,
    dated: batch.dated,
    downloadStartDate: batch.downloadStartDate,
    downloadEndDate: batch.downloadEndDate,
    work
  }
}

function rowKey(batchId: string, serial: number): string {
  return `${batchId}-${serial}`
}

function suggestedName(work: BidDocumentWork): string {
  // Save each Bid Document under its work type + location + estimate, e.g.
  // "Bid Document: CC Roads @ Venkateshwaranagar 45 Lakhs". The Works
  // List's own "Type of work" column wins when the row has one; a blank
  // column falls back to inferTypeOfWork(name) via resolveTypeLabel (same
  // rule the Tender Notice's own page-3 count breakdown uses — see
  // core/tenderNotice.ts). Location is a best-effort guess off the work
  // name (core/workLocationExtractor.ts) and is left out of the name
  // entirely when nothing plausible was found, rather than showing a
  // misleading placeholder. The colon is safe even though Windows
  // disallows it in file names — sanitizeFileName (electron/main.ts)
  // strips it to "-" before the name hits disk.
  const typeLabel = resolveTypeLabel({ name: work.name, amount: work.amount, typeOfWork: work.typeOfWork })
  const location = inferWorkLocation(work.name)
  const locationTag = location ? ` @ ${location}` : ''
  return `Bid Document: ${typeLabel}${locationTag} ${work.amount} Lakhs`.slice(0, 150)
}

/**
 * Bid Documents to generate for each tender notice issued from the Calendar
 * section — one "BID Document N" per work in the notice's item table, N
 * matching that work's serial number there. Each row fills the bundled Bid
 * Document template with that work's own Name/Estimate/Zone/Circle/
 * Completion Period/EMD plus the notice's NIT No. and dates.
 */
export default function BidDocumentsPanel({ batches, onRemove, onUpdateWork }: Props) {
  const [busyKey, setBusyKey] = useState<string | null>(null)
  const [savedKey, setSavedKey] = useState<string | null>(null)
  const [batchBusyId, setBatchBusyId] = useState<string | null>(null)
  const [batchProgress, setBatchProgress] = useState<BidDocumentBatchProgress | null>(null)
  const [batchSavedDir, setBatchSavedDir] = useState<{ batchId: string; dir: string } | null>(null)
  const [error, setError] = useState<string | null>(null)
  // Each batch's work list auto-collapses on arrival (a Calendar session can
  // pile up several NIT batches, each with dozens of works) — a batch id
  // only stays expanded once the user explicitly opens it here.
  const [expandedBatches, setExpandedBatches] = useState<Set<string>>(new Set())
  const toggleExpanded = (id: string) =>
    setExpandedBatches((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })
  const [previewing, setPreviewing] = useState<{ batch: BidDocumentBatch; work: BidDocumentWork } | null>(
    null
  )
  const previewRef = useRef<HTMLDivElement>(null)
  const batchBusyRef = useRef(batchBusyId)
  batchBusyRef.current = batchBusyId

  useEffect(() => {
    return api.onBidDocumentBatchProgress((p) => {
      // Ignore a straggling event from a batch that's no longer active.
      if (batchBusyRef.current) setBatchProgress(p)
    })
  }, [])

  if (batches.length === 0) return null

  async function preview(batch: BidDocumentBatch, work: BidDocumentWork) {
    const key = rowKey(batch.id, work.serial)
    setError(null)
    setBusyKey(key)
    try {
      const b64 = await api.previewBidDocument(buildInput(batch, work))
      setPreviewing({ batch, work })
      requestAnimationFrame(() => {
        void (async () => {
          const container = previewRef.current
          if (!container) return
          await renderDocPreview(base64ToUint8(b64), container)
        })()
      })
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusyKey(null)
    }
  }

  async function save(batch: BidDocumentBatch, work: BidDocumentWork) {
    const key = rowKey(batch.id, work.serial)
    setError(null)
    setBusyKey(key)
    setSavedKey(null)
    try {
      const path = await api.generateBidDocument(buildInput(batch, work), suggestedName(work))
      if (path) setSavedKey(key)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusyKey(null)
    }
  }

  async function downloadAll(batch: BidDocumentBatch) {
    setError(null)
    setBatchBusyId(batch.id)
    setBatchProgress(null)
    setBatchSavedDir(null)
    try {
      const paths = await api.generateBidDocumentBatch(
        batch.works.map((work) => ({ input: buildInput(batch, work), suggestedName: suggestedName(work) }))
      )
      if (paths && paths.length > 0) {
        setBatchSavedDir({ batchId: batch.id, dir: paths[0].replace(/[^/\\]+$/, '') })
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBatchProgress(null)
      setBatchBusyId(null)
    }
  }

  const batchPct =
    batchBusyId && batchProgress && batchProgress.total > 0
      ? Math.round((batchProgress.done / batchProgress.total) * 100)
      : null

  return (
    <>
      {batches.map((batch) => {
        const expanded = expandedBatches.has(batch.id)
        return (
        <section className="card bid-batch" key={batch.id}>
          <div className="card-head">
            <button
              className="ghost bid-batch-toggle"
              title={expanded ? 'Collapse' : 'Expand'}
              onClick={() => toggleExpanded(batch.id)}
            >
              <IconChevronRight className={expanded ? 'bid-batch-toggle-icon open' : 'bid-batch-toggle-icon'} />
            </button>
            <div className="head-ic">
              <IconDoc />
            </div>
            <div className="titles">
              <h2>Bid Documents</h2>
              <p className="sub">NIT No. {batch.nitNo}</p>
            </div>
            <button
              className="primary scan-progress-btn"
              onClick={() => downloadAll(batch)}
              disabled={batchBusyId === batch.id}
            >
              {batchBusyId === batch.id && batchPct !== null && (
                <span className="scan-progress-fill" style={{ width: `${batchPct}%` }} />
              )}
              <span className="scan-progress-label">
                <IconDownload />{' '}
                {batchBusyId === batch.id
                  ? batchPct !== null
                    ? `Saving ${batchProgress!.done}/${batchProgress!.total} (${batchPct}%)`
                    : 'Saving…'
                  : `Download All (${batch.works.length})`}
              </span>
            </button>
            <button className="danger-ghost" title="Remove this batch" onClick={() => onRemove(batch.id)}>
              <IconTrash />
            </button>
          </div>

          {batchSavedDir?.batchId === batch.id && (
            <div className="notice ok">
              <IconCheck />
              Saved {batch.works.length} Bid Document{batch.works.length === 1 ? '' : 's'} to {batchSavedDir.dir}
            </div>
          )}

          {expanded && (
          <ul className="todo-list boq-entries">
            {batch.works.map((work) => {
              const key = rowKey(batch.id, work.serial)
              const se = isSeWork(work)
              const set = (patch: Partial<BidDocumentWork>) => onUpdateWork(batch.id, work.serial, patch)
              return (
                <li key={key} className={`todo-item ${se ? 'bid-se-row' : ''}`}>
                  <div className="todo-body">
                    <span className="todo-text">
                      BID Document {work.serial} {se && <span className="badge info">SE office</span>}
                    </span>
                    <span className="todo-dates">{work.name || 'Unnamed work'}</span>
                    {se && (
                      <div className="tender-field-row bid-se-fields">
                        <label className="tender-field">
                          <span>Item No.</span>
                          <input
                            value={work.itemNo ?? ''}
                            onChange={(e) => set({ itemNo: e.target.value })}
                            placeholder={String(work.serial)}
                          />
                        </label>
                        <label className="tender-field">
                          <span>T.S. No.</span>
                          <input
                            value={work.tsNo ?? ''}
                            onChange={(e) => set({ tsNo: e.target.value })}
                            placeholder="29/SE/QBZ/CMC/2026-27"
                          />
                        </label>
                        <label className="tender-field">
                          <span>T.S. Date</span>
                          <input
                            value={work.tsDate ?? ''}
                            onChange={(e) => set({ tsDate: e.target.value })}
                            placeholder="04-07-2026"
                          />
                        </label>
                        <label className="tender-field">
                          <span>AS Authority</span>
                          <select
                            value={work.asAuthority ?? 'commissioner'}
                            onChange={(e) => set({ asAuthority: e.target.value as 'zonal' | 'commissioner' | 'resolution' })}
                          >
                            <option value="commissioner">Commissioner, CMC</option>
                            <option value="zonal">Zonal Commissioner</option>
                            <option value="resolution">Standing committee / resolution</option>
                          </select>
                        </label>
                        <label className="tender-field">
                          <span>AS Date</span>
                          <input
                            value={work.asDate ?? ''}
                            onChange={(e) => set({ asDate: e.target.value })}
                            placeholder="23.06.2026"
                          />
                        </label>
                      </div>
                    )}
                  </div>
                  <button
                    className="ghost"
                    title="Preview"
                    disabled={busyKey === key}
                    onClick={() => preview(batch, work)}
                  >
                    <IconEye />
                  </button>
                  <button
                    className="ghost"
                    title="Save"
                    disabled={busyKey === key}
                    onClick={() => save(batch, work)}
                  >
                    {savedKey === key ? <IconCheck /> : <IconDownload />}
                  </button>
                </li>
              )
            })}
          </ul>
          )}

          {error && (
            <div className="notice error">
              <IconWarn />
              {error}
            </div>
          )}
        </section>
        )
      })}

      {previewing &&
        createPortal(
          <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => setPreviewing(null))}>
            <div className="tender-modal tender-modal-wide" onClick={(e) => e.stopPropagation()}>
              <h3>
                BID Document {previewing.work.serial} — {previewing.work.name || 'preview'}
              </h3>
              <div className="tender-preview-scroll">
                <div ref={previewRef} className="docprev-body" />
              </div>
              <div className="confirm-actions">
                <button className="ghost" onClick={() => setPreviewing(null)}>
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

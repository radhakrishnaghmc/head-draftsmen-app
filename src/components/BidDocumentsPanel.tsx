import { useRef, useState } from 'react'
import { api } from '../ipc'
import { IconDoc, IconEye, IconDownload, IconTrash, IconCheck, IconWarn } from './Icons'
import { base64ToUint8, renderDocPreview } from './docPage'
import type { BidDocumentBatch, BidDocumentWork } from '@core/types'
import type { BidDocumentInput } from '../../electron/ipc-contract'
import { closeOnBackdropMouseDown } from '../overlayClose'

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
  // Save each Bid Document under its Win Code, e.g. "Bid Document 12345".
  // Falls back to the serial-number label for older batches saved before the
  // Win Code was carried through.
  const id = work.winCode?.trim() || String(work.serial)
  return `Bid Document ${id}`.slice(0, 150)
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
  const [batchSavedDir, setBatchSavedDir] = useState<{ batchId: string; dir: string } | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [previewing, setPreviewing] = useState<{ batch: BidDocumentBatch; work: BidDocumentWork } | null>(
    null
  )
  const previewRef = useRef<HTMLDivElement>(null)

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
      setBatchBusyId(null)
    }
  }

  return (
    <>
      {batches.map((batch) => (
        <section className="card bid-batch" key={batch.id}>
          <div className="card-head">
            <div className="head-ic">
              <IconDoc />
            </div>
            <div className="titles">
              <h2>Bid Documents</h2>
              <p className="sub">NIT No. {batch.nitNo}</p>
            </div>
            <button className="primary" onClick={() => downloadAll(batch)} disabled={batchBusyId === batch.id}>
              <IconDownload /> {batchBusyId === batch.id ? 'Saving…' : `Download All (${batch.works.length})`}
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
                            onChange={(e) => set({ asAuthority: e.target.value as 'zonal' | 'commissioner' })}
                          >
                            <option value="commissioner">Commissioner, CMC</option>
                            <option value="zonal">Zonal Commissioner</option>
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

          {error && (
            <div className="notice error">
              <IconWarn />
              {error}
            </div>
          )}
        </section>
      ))}

      {previewing && (
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
        </div>
      )}
    </>
  )
}

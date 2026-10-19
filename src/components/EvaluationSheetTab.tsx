import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfToTextLinesOrOcr } from '../pdfToText'
import { readEvaluationSheetPage } from '@core/tenderAgents/evaluationSheetAgent'
import { BID_CAPACITY_THRESHOLD, defaultConditions, isBidCapacityCondition, type EvaluationSheetInput } from '@core/evaluationSheet'
import type { Office } from '../office'
import type { ExcelTable, TenderConditionsEntry } from '@core/types'
import { matchWorksRow } from '../estimateDownloads'
import { lakhsToRupees } from '@core/worksAmounts'
import { IconFolder, IconTable, IconPlus, IconTrash, IconChecklist, IconWarn, IconDownload } from './Icons'

interface Props {
  office: Office
  /** Works List tables; the first is the office's Works List. */
  tables: ExcelTable[]
  /** Tender conditions saved in the app by tender ID (the portal drops a tender once bidding closes). */
  savedConditions: Record<string, TenderConditionsEntry>
  onSaveConditions: (tenderId: string, conditions: string[]) => void
  /** Called once sheets are saved, so those tenders' conditions can be dropped a week later. */
  onEvaluationGenerated: (tenderIds: string[]) => void
}

interface SheetEntry {
  id: string
  sourceName: string
  tenderId: string
  nitLine: string
  workLine: string
  ecv: string
  /** Estimate amount in rupees from the Works List ("" when the work isn't found there). */
  estimate: string
  /** Where the estimate came from / why it is missing. */
  estimateNote: string | null
  signature: string
  bidders: string[]
  /** The conditions printed on the sheet (tender's own list or the standard ones), editable. */
  conditions: string[]
  /** Why the portal conditions could not be used (shown as a hint). */
  conditionsNote: string | null
  error: string | null
  saved: string | null
  /** The ECV the small-work question was last answered for, so it isn't asked again for the same value. */
  askedEcv: string | null
}

function newId(): string {
  return `ev-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader()
    r.onload = () => resolve(r.result as string)
    r.onerror = () => reject(r.error)
    r.readAsDataURL(file)
  })
}

function suggestedName(e: SheetEntry): string {
  return `Evaluation Sheet${e.tenderId ? ` ${e.tenderId}` : ''}`
}

/**
 * Evaluation Sheet — upload the portal's "View Bidders" (Supplier List) PDFs,
 * or phone photos of the same page, and issue one Bid Capacity Evaluation
 * Sheet (.xlsx) per tender: one column per participating bidder, the eleven
 * fixed tender-condition rows, and the 3AN−B bid-capacity block wired as
 * formulas. The header, bidder columns and ECV are filled from each upload;
 * the per-bidder turnover / N / B / eligibility cells are left blank for the
 * office to key in from each bidder's documents. Several tenders can be
 * loaded at once and their sheets downloaded together into one folder.
 */
export default function EvaluationSheetTab({ office, tables, savedConditions, onSaveConditions, onEvaluationGenerated }: Props) {
  const [entries, setEntries] = useState<SheetEntry[]>([])
  const [busy, setBusy] = useState<null | 'reading' | 'downloadAll'>(null)
  // Upload progress shown inside the upload button: file `index` of `total`, in `phase`.
  const [progress, setProgress] = useState<{ index: number; total: number; phase: 'reading' | 'estimate' | 'conditions' } | null>(null)
  const [pickError, setPickError] = useState<string | null>(null)
  const [savedAllDir, setSavedAllDir] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  // Manual tender ID: fetch that tender's conditions from the portal and save them by ID.
  const [manualId, setManualId] = useState('')
  const [manualBusy, setManualBusy] = useState(false)
  const [manualMsg, setManualMsg] = useState<{ ok: boolean; text: string } | null>(null)

  async function saveManualConditions() {
    const id = manualId.trim()
    if (!id) return
    if (!/^\d+$/.test(id)) {
      setManualMsg({ ok: false, text: 'Enter the numeric tender ID.' })
      return
    }
    if (savedConditions[id]?.conditions.length) {
      setManualMsg({ ok: true, text: `Tender ${id} conditions are already saved (${savedConditions[id].conditions.length}).` })
      return
    }
    setManualBusy(true)
    setManualMsg(null)
    try {
      const conditions = await api.fetchTenderConditions(id)
      onSaveConditions(id, conditions)
      setManualMsg({ ok: true, text: `Saved ${conditions.length} tender conditions for tender ${id}. Its evaluation sheet will use them.` })
      // Sheets already loaded for this tender switch to the saved conditions.
      setEntries((prev) =>
        prev.map((x) =>
          x.tenderId === id
            ? {
                ...x,
                conditions: defaultConditions(conditions, x.workLine, Number(x.estimate.replace(/,/g, '')) || null),
                conditionsNote: `Using the tender conditions saved earlier for tender ${id}.`
              }
            : x
        )
      )
    } catch (err) {
      const why = err instanceof Error ? err.message.replace(/^Error invoking remote method '[^']*': (Error: )?/, '') : String(err)
      setManualMsg({ ok: false, text: `Could not save the conditions (${why}). Standard conditions are used unless some were saved earlier.` })
    } finally {
      setManualBusy(false)
    }
  }
  // Pop-up asking whether a work missing from the Works List has an estimate of 5 lakhs or below.
  const [estimateAsk, setEstimateAsk] = useState<{ work: string; why: string; answer: (below: boolean) => void } | null>(null)

  function askEstimateBelowFiveLakhs(work: string, why: string): Promise<boolean> {
    return new Promise((resolve) =>
      setEstimateAsk({
        work,
        why,
        answer: (below) => {
          setEstimateAsk(null)
          resolve(below)
        }
      })
    )
  }

  function defaultSignature(): string {
    const circle = (office.circle ?? '').trim()
    const cno = (office.circleNumber ?? '').trim()
    const corp = (office.corporation ?? '').trim() || 'CMC'
    const circlePart = circle ? `${circle} circle${cno ? `-${cno}` : ''},` : ''
    return `Executive Engineer,\n${circlePart}${corp}`
  }

  async function handleFiles(files: File[]) {
    setBusy('reading')
    setPickError(null)
    setSavedAllDir(null)
    try {
      for (const [fileIndex, file] of files.entries()) {
        setProgress({ index: fileIndex, total: files.length, phase: 'reading' })
        const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
        try {
          const lines = isPdf ? await pdfToTextLinesOrOcr(file) : await api.ocrPhotosToLines([await readAsDataUrl(file)])
          // Evaluation Sheet agent: header fields + participating bidders from the page.
          const { header: ev, bidders: list } = readEvaluationSheetPage(lines)
          // Estimate amount from the Works List, matched on the work's name — it
          // decides whether the sheet needs the 3AN−B rows (estimate > 5 lakhs).
          let estimate = ''
          let estimateNote: string | null = null
          const worksTable = tables[0]
          if (!ev.nameOfWork) {
            estimateNote = 'No work name found in the upload, so the Works List estimate could not be looked up.'
          } else if (!worksTable) {
            estimateNote = 'No Works List loaded, so the estimate could not be looked up.'
          } else {
            setProgress({ index: fileIndex, total: files.length, phase: 'estimate' })
            const match = await matchWorksRow(ev.nameOfWork, worksTable)
            const rupees = match ? lakhsToRupees(match.row['Amount of estimate'] ?? '') : 0
            if (rupees > 0) {
              estimate = String(rupees)
              estimateNote = `Estimate ₹${rupees.toLocaleString('en-IN')} from the Works List${match?.matchedViaAi ? ' (AI name match — please verify)' : ''}.`
            } else {
              estimateNote = `${match ? 'The Works List row has no estimate amount' : 'Work not found on the Works List'}.`
            }
          }
          // The estimate is always above the ECV, so an ECV over 5 lakhs already means
          // 3AN−B applies. A smaller ECV leaves the question open whatever the Works
          // List says, so ask: yes drops the bid-capacity conditions and rows.
          if (ev.nameOfWork && ev.ecvRupees != null && ev.ecvRupees < BID_CAPACITY_THRESHOLD) {
            const below = await askEstimateBelowFiveLakhs(ev.nameOfWork, `The ECV is ₹${ev.ecvRupees.toLocaleString('en-IN')}, below ₹5 lakhs`)
            if (below) {
              estimate = String(BID_CAPACITY_THRESHOLD)
              estimateNote = 'Estimate confirmed below ₹5 lakhs, so the 3AN−B conditions and rows are left out.'
            } else {
              estimateNote = `${estimateNote ? estimateNote + ' ' : ''}Estimate confirmed ₹5 lakhs or above, so the 3AN−B rows are kept.`
            }
          }
          // The View Bidders sheet names the tender; look it up on the portal and
          // take its listed conditions in place of the standard eleven.
          let fetched: string[] | null = null
          let conditionsNote: string | null = null
          if (ev.tenderId) {
            setProgress({ index: fileIndex, total: files.length, phase: 'conditions' })
            // Conditions saved earlier win: a closed tender is gone from the portal.
            fetched = savedConditions[ev.tenderId]?.conditions.length ? savedConditions[ev.tenderId].conditions : null
            if (fetched) {
              conditionsNote = `Using the tender conditions saved earlier for tender ${ev.tenderId}.`
            } else {
              try {
                fetched = await api.fetchTenderConditions(ev.tenderId)
                onSaveConditions(ev.tenderId, fetched)
              } catch (err) {
                const why = err instanceof Error ? err.message.replace(/^Error invoking remote method '[^']*': (Error: )?/, '') : String(err)
                conditionsNote = `Could not fetch the tender conditions from the portal (${why}); using the standard conditions. Enter this tender ID above while the tender is still open to save its conditions.`
              }
            }
          } else {
            conditionsNote = 'No tender ID found in the upload, so the standard conditions are used.'
          }
          const entry: SheetEntry = {
            id: newId(),
            sourceName: file.name,
            tenderId: ev.tenderId ?? '',
            nitLine: `NIT No: ${ev.noticeNo ?? ''}${ev.noticeDate ? `,Dt:${ev.noticeDate}` : ''}${ev.tenderId ? `,Tender ID:${ev.tenderId}` : ''}`,
            workLine: `Name of the work: ${ev.nameOfWork ?? ''}${ev.ecvRupees != null ? `. ecv: ${ev.ecvRupees.toFixed(2)}` : ''}`,
            ecv: ev.ecvRupees != null ? String(ev.ecvRupees) : '',
            signature: defaultSignature(),
            bidders: list.map((b) => b.name),
            estimate,
            estimateNote,
            conditions: defaultConditions(fetched, `Name of the work: ${ev.nameOfWork ?? ''}`, estimate ? Number(estimate) || null : null),
            conditionsNote,
            error:
              list.length === 0
                ? `No participating bidders found in ${file.name}. Upload the portal's "View Bidders" (Supplier List) page, as a PDF or a clear photo.`
                : null,
            saved: null,
            askedEcv: ev.ecvRupees != null && ev.ecvRupees < BID_CAPACITY_THRESHOLD ? String(ev.ecvRupees) : null
          }
          setEntries((prev) => [...prev, entry])
        } catch (e) {
          setEntries((prev) => [
            ...prev,
            {
              id: newId(),
              sourceName: file.name,
              tenderId: '',
              nitLine: '',
              workLine: '',
              ecv: '',
              signature: defaultSignature(),
              bidders: [],
              estimate: '',
              estimateNote: null,
              conditions: [],
              conditionsNote: null,
              error: e instanceof Error ? e.message : String(e),
              saved: null,
              askedEcv: null
            }
          ])
        }
      }
    } finally {
      setBusy(null)
      setProgress(null)
    }
  }

  function removeEntry(id: string) {
    setEntries((prev) => prev.filter((e) => e.id !== id))
  }

  function updateEntry(id: string, patch: Partial<SheetEntry>) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, ...patch } : e)))
  }

  function setCondition(id: string, i: number, value: string) {
    setEntries((prev) => prev.map((x) => (x.id === id ? { ...x, conditions: x.conditions.map((c, j) => (j === i ? value : c)) } : x)))
  }
  function addCondition(id: string) {
    setEntries((prev) => prev.map((x) => (x.id === id ? { ...x, conditions: [...x.conditions, ''] } : x)))
  }
  function removeCondition(id: string, i: number) {
    setEntries((prev) => prev.map((x) => (x.id === id ? { ...x, conditions: x.conditions.filter((_, j) => j !== i) } : x)))
  }

  function setBidder(id: string, i: number, value: string) {
    setEntries((prev) =>
      prev.map((e) => (e.id === id ? { ...e, bidders: e.bidders.map((b, j) => (j === i ? value : b)) } : e))
    )
  }
  function addBidder(id: string) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, bidders: [...e.bidders, ''] } : e)))
  }
  function removeBidder(id: string, i: number) {
    setEntries((prev) =>
      prev.map((e) => (e.id === id ? { ...e, bidders: e.bidders.filter((_, j) => j !== i) } : e))
    )
  }

  function toInput(e: SheetEntry): EvaluationSheetInput | null {
    const bidders = e.bidders.map((b) => b.trim()).filter(Boolean)
    if (bidders.length === 0) return null
    return { nitLine: e.nitLine, workLine: e.workLine, bidders, ecv: Number(e.ecv.replace(/,/g, '')) || null, estimate: Number(e.estimate.replace(/,/g, '')) || null, signature: e.signature, conditions: e.conditions, conditionsFinal: true }
  }

  /**
   * The ECV can be edited after upload. Below ₹5 lakhs, ask once per ECV value
   * whether the estimate is too; yes drops the bid-capacity conditions and rows.
   * Returns the entry to download (updated in place on screen too).
   */
  async function confirmSmallWork(e: SheetEntry): Promise<SheetEntry> {
    const ecv = Number(e.ecv.replace(/,/g, '')) || 0
    if (ecv <= 0 || ecv >= BID_CAPACITY_THRESHOLD || e.askedEcv === e.ecv) return e
    const below = await askEstimateBelowFiveLakhs(e.workLine.replace(/^Name of the work:\s*/i, ''), `The ECV is ₹${ecv.toLocaleString('en-IN')}, below ₹5 lakhs`)
    const patch: Partial<SheetEntry> = below
      ? {
          askedEcv: e.ecv,
          estimate: String(BID_CAPACITY_THRESHOLD),
          estimateNote: 'Estimate confirmed below ₹5 lakhs, so the 3AN−B conditions and rows are left out.',
          conditions: e.conditions.filter((c) => !isBidCapacityCondition(c))
        }
      : { askedEcv: e.ecv, estimateNote: 'Estimate confirmed ₹5 lakhs or above, so the 3AN−B rows are kept.' }
    updateEntry(e.id, patch)
    return { ...e, ...patch }
  }

  async function downloadOne(entry: SheetEntry) {
    const e = await confirmSmallWork(entry)
    const input = toInput(e)
    if (!input) {
      updateEntry(e.id, { error: 'Add at least one bidder before downloading.' })
      return
    }
    setBusy('reading')
    updateEntry(e.id, { error: null, saved: null })
    try {
      const path = await api.exportEvaluationSheet(input, suggestedName(e))
      updateEntry(e.id, { saved: path ? `Saved to ${path}` : 'Cancelled.' })
      if (path && e.tenderId) onEvaluationGenerated([e.tenderId])
    } catch (err) {
      updateEntry(e.id, { error: err instanceof Error ? err.message : String(err) })
    } finally {
      setBusy(null)
    }
  }

  async function downloadAll() {
    const confirmed: SheetEntry[] = []
    for (const entry of entries) confirmed.push(entry.bidders.some((b) => b.trim()) ? await confirmSmallWork(entry) : entry)
    const ready = confirmed.map((e) => ({ e, input: toInput(e) })).filter((x) => x.input != null)
    if (ready.length === 0) {
      setPickError('Add at least one bidder before downloading.')
      return
    }
    setBusy('downloadAll')
    setPickError(null)
    setSavedAllDir(null)
    try {
      const paths = await api.exportEvaluationSheetBatch(
        ready.map(({ e, input }) => ({ input: input as EvaluationSheetInput, suggestedName: suggestedName(e) }))
      )
      if (paths && paths.length > 0) {
        setSavedAllDir(paths[0].replace(/[^/\\]+$/, ''))
        onEvaluationGenerated(ready.map(({ e }) => e.tenderId).filter(Boolean))
      }
    } catch (e) {
      setPickError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // Each file counts as three steps: read the PDF, look up its estimate, fetch its conditions.
  const pct = progress
    ? Math.round(((progress.index + (progress.phase === 'conditions' ? 0.66 : progress.phase === 'estimate' ? 0.33 : 0)) / progress.total) * 100)
    : 0
  const readyCount = entries.filter((e) => e.bidders.some((b) => b.trim())).length

  return (
    <div className="card">
      {estimateAsk && (
        <div className="wo-modal-overlay">
          <div className="wo-modal office-modal" role="dialog" aria-modal="true" style={{ padding: 20 }}>
            <h3 style={{ marginTop: 0 }}>Estimate amount</h3>
            <p className="estimate-hint">{estimateAsk.why}:</p>
            <p>{estimateAsk.work}</p>
            <p>Is the Estimate amount less than 5 Lakhs? (3AN−B is checked only at ₹5 lakhs and above.)</p>
            <div className="boq-actions">
              <button className="primary" onClick={() => estimateAsk.answer(true)}>
                Yes, less than 5 lakhs
              </button>
              <button className="ghost" onClick={() => estimateAsk.answer(false)}>
                No, 5 lakhs or above
              </button>
            </div>
          </div>
        </div>
      )}
      <div className="empty empty--tight">
        <IconChecklist />
        <div className="boq-actions">
          <button
            className="primary upload-btn"
            onClick={() => fileInputRef.current?.click()}
            disabled={busy !== null}
            style={
              progress
                ? { background: `linear-gradient(to right, rgba(255,255,255,0.28) ${pct}%, transparent ${pct}%), var(--primary, #2563eb)` }
                : undefined
            }
          >
            <IconFolder />{' '}
            {progress
              ? `${progress.phase === 'reading' ? 'Reading' : progress.phase === 'estimate' ? 'Looking up estimate' : 'Fetching tender conditions'} ${progress.index + 1}/${progress.total}… ${pct}%`
              : busy === 'reading'
                ? 'Reading…'
                : entries.length > 0
                  ? 'Add more tenders'
                  : 'Upload View Bidders PDFs / photos'}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept="application/pdf,.pdf,image/*"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => {
              const files = Array.from(e.target.files ?? [])
              if (files.length > 0) void handleFiles(files)
              e.target.value = ''
            }}
          />
          {readyCount > 1 && (
            <button className="ghost upload-btn" onClick={downloadAll} disabled={busy !== null}>
              <IconDownload /> {busy === 'downloadAll' ? 'Saving…' : `Download all ${readyCount} evaluation sheets`}
            </button>
          )}
        </div>
        <div className="boq-actions">
          <input
            value={manualId}
            onChange={(ev) => setManualId(ev.target.value)}
            onKeyDown={(ev) => {
              if (ev.key === 'Enter') void saveManualConditions()
            }}
            onBlur={() => void saveManualConditions()}
            disabled={manualBusy}
            placeholder={manualBusy ? 'Saving conditions…' : 'Tender ID — press Enter to save its conditions'}
            inputMode="numeric"
            aria-label="Tender ID"
            style={{ maxWidth: 360 }}
          />
        </div>
        {manualMsg && (
          <div className={`notice ${manualMsg.ok ? 'ok' : 'error'}`}>
            {manualMsg.ok ? <IconTable /> : <IconWarn />} {manualMsg.text}
          </div>
        )}
        {pickError && (
          <div className="notice error">
            <IconWarn /> {pickError}
          </div>
        )}
        {savedAllDir && (
          <div className="notice ok">
            <IconTable /> Saved {readyCount} evaluation sheets to {savedAllDir}
          </div>
        )}
      </div>

      {entries.length > 0 && (
        <ul className="estimate-list">
          {entries.map((e) => (
            <li key={e.id} className="estimate-entry">
              <div className="estimate-entry-head">
                <div className="estimate-entry-titles">
                  <span className="estimate-entry-file">{e.sourceName}</span>
                  {e.workLine && <h3 className="estimate-work-name">{e.workLine.replace(/^Name of the work:\s*/, '')}</h3>}
                </div>
                <button className="danger-ghost" title="Remove" onClick={() => removeEntry(e.id)}>
                  <IconTrash />
                </button>
              </div>

              {e.error && (
                <p className="estimate-entry-error">
                  <IconWarn /> {e.error}
                </p>
              )}
              {e.saved && <p className="estimate-hint">{e.saved}</p>}
              {e.estimateNote && <p className="estimate-hint">{e.estimateNote}</p>}
              {e.conditionsNote && <p className="estimate-hint">{e.conditionsNote}</p>}
              {e.bidders.length > 0 && (
                <div className="estimate-body">
                  <div className="loa-manual-fields">
                    <label className="field-label">
                      NIT / Tender line
                      <input value={e.nitLine} onChange={(ev) => updateEntry(e.id, { nitLine: ev.target.value })} />
                    </label>
                    <label className="field-label">
                      Name of work line
                      <input value={e.workLine} onChange={(ev) => updateEntry(e.id, { workLine: ev.target.value })} />
                    </label>
                    <div className="loa-manual-grid">
                      <label className="field-label">
                        ECV (rupees)
                        <input value={e.ecv} onChange={(ev) => updateEntry(e.id, { ecv: ev.target.value })} />
                      </label>
                      <label className="field-label">
                        Signature
                        <input
                          value={e.signature.replace(/\n/g, ' — ')}
                          onChange={(ev) => updateEntry(e.id, { signature: ev.target.value.replace(/ — /g, '\n') })}
                        />
                      </label>
                    </div>

                    <div className="ns-group-title">
                      Tender conditions ({e.conditions.length})
                      <button className="ns-add" onClick={() => addCondition(e.id)} title="Add a condition">
                        <IconPlus /> Add
                      </button>
                    </div>
                    <div className="ns-bidders">
                      {e.conditions.map((c, i) => (
                        <div className="ns-bidder-row" key={i} style={{ gridTemplateColumns: '2rem 1fr 2rem' }}>
                          <span className="mb-serial">{i + 1}</span>
                          <textarea
                            value={c}
                            rows={Math.min(5, Math.max(1, Math.ceil(c.length / 70)))}
                            onChange={(ev) => setCondition(e.id, i, ev.target.value)}
                            placeholder="Tender condition"
                            style={{ resize: 'vertical', font: 'inherit' }}
                          />
                          <button className="ns-del" onClick={() => removeCondition(e.id, i)} title="Remove">
                            ✕
                          </button>
                        </div>
                      ))}
                    </div>

                    <div className="ns-group-title">
                      Participating bidders ({e.bidders.length})
                      <button className="ns-add" onClick={() => addBidder(e.id)} title="Add a bidder">
                        <IconPlus /> Add
                      </button>
                    </div>
                    <div className="ns-bidders">
                      {e.bidders.map((b, i) => (
                        <div className="ns-bidder-row" key={i} style={{ gridTemplateColumns: '2rem 1fr 2rem' }}>
                          <span className="mb-serial">{i + 1}</span>
                          <input value={b} onChange={(ev) => setBidder(e.id, i, ev.target.value)} placeholder="Bidder / firm name" />
                          <button className="ns-del" onClick={() => removeBidder(e.id, i)} title="Remove">
                            ✕
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="doc-sheet-footer">
                    <span className="estimate-hint">
                      Fills the header, one column per bidder and the ECV, and wires the 3AN−B bid-capacity formula.
                      Turnover (A), N, B, the eligibility rows and Remarks are left blank to fill in from each bidder's
                      documents.
                    </span>
                    <button className="primary" onClick={() => downloadOne(e)} disabled={busy !== null}>
                      <IconTable /> Download evaluation sheet
                    </button>
                  </div>
                </div>
              )}
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

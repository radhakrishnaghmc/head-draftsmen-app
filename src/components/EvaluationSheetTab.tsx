import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfToTextLines } from '../pdfToText'
import { parseTenderEvaluation } from '@core/tenderEvaluationPdf'
import { parseParticipatingBidders } from '@core/viewBiddersPdf'
import type { EvaluationSheetInput } from '@core/evaluationSheet'
import type { Office } from '../office'
import { IconFolder, IconTable, IconPlus, IconTrash, IconChecklist, IconWarn, IconDownload } from './Icons'

interface Props {
  office: Office
}

interface SheetEntry {
  id: string
  sourceName: string
  tenderId: string
  nitLine: string
  workLine: string
  ecv: string
  signature: string
  bidders: string[]
  error: string | null
  saved: string | null
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
  return `Evaluation Sheet${e.tenderId ? ` - ${e.tenderId}` : ''}`
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
export default function EvaluationSheetTab({ office }: Props) {
  const [entries, setEntries] = useState<SheetEntry[]>([])
  const [busy, setBusy] = useState<null | 'reading' | 'downloadAll'>(null)
  const [pickError, setPickError] = useState<string | null>(null)
  const [savedAllDir, setSavedAllDir] = useState<string | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)

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
      for (const file of files) {
        const isPdf = /\.pdf$/i.test(file.name) || file.type === 'application/pdf'
        try {
          const lines = isPdf ? await pdfToTextLines(file) : await api.ocrPhotosToLines([await readAsDataUrl(file)])
          const ev = parseTenderEvaluation(lines)
          const list = parseParticipatingBidders(lines)
          const entry: SheetEntry = {
            id: newId(),
            sourceName: file.name,
            tenderId: ev.tenderId ?? '',
            nitLine: `NIT No: ${ev.noticeNo ?? ''}${ev.noticeDate ? `,Dt:${ev.noticeDate}` : ''}${ev.tenderId ? `,Tender ID:${ev.tenderId}` : ''}`,
            workLine: `Name of the work: ${ev.nameOfWork ?? ''}${ev.ecvRupees != null ? `. ecv: ${ev.ecvRupees.toFixed(2)}` : ''}`,
            ecv: ev.ecvRupees != null ? String(ev.ecvRupees) : '',
            signature: defaultSignature(),
            bidders: list.map((b) => b.name),
            error:
              list.length === 0
                ? `No participating bidders found in ${file.name}. Upload the portal's "View Bidders" (Supplier List) page, as a PDF or a clear photo.`
                : null,
            saved: null
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
              error: e instanceof Error ? e.message : String(e),
              saved: null
            }
          ])
        }
      }
    } finally {
      setBusy(null)
    }
  }

  function removeEntry(id: string) {
    setEntries((prev) => prev.filter((e) => e.id !== id))
  }

  function updateEntry(id: string, patch: Partial<SheetEntry>) {
    setEntries((prev) => prev.map((e) => (e.id === id ? { ...e, ...patch } : e)))
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
    return { nitLine: e.nitLine, workLine: e.workLine, bidders, ecv: Number(e.ecv.replace(/,/g, '')) || null, signature: e.signature }
  }

  async function downloadOne(e: SheetEntry) {
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
    } catch (err) {
      updateEntry(e.id, { error: err instanceof Error ? err.message : String(err) })
    } finally {
      setBusy(null)
    }
  }

  async function downloadAll() {
    const ready = entries.map((e) => ({ e, input: toInput(e) })).filter((x) => x.input != null)
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
      if (paths && paths.length > 0) setSavedAllDir(paths[0].replace(/[^/\\]+$/, ''))
    } catch (e) {
      setPickError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  const readyCount = entries.filter((e) => e.bidders.some((b) => b.trim())).length

  return (
    <div className="card">
      <div className="empty empty--tight">
        <IconChecklist />
        <div className="boq-actions">
          <button className="primary upload-btn" onClick={() => fileInputRef.current?.click()} disabled={busy !== null}>
            <IconFolder /> {busy === 'reading' ? 'Reading…' : entries.length > 0 ? 'Add more tenders' : 'Upload View Bidders PDFs / photos'}
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

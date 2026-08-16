import { useRef, useState } from 'react'
import { api } from '../ipc'
import { pdfToTextLines } from '../pdfToText'
import { parseTenderEvaluation } from '@core/tenderEvaluationPdf'
import { parseParticipatingBidders } from '@core/viewBiddersPdf'
import type { Office } from '../office'
import { IconFolder, IconTable, IconPlus, IconTrash, IconChecklist } from './Icons'

interface Props {
  office: Office
}

/**
 * Evaluation Sheet — upload the portal's "View Bidders" (Supplier List) PDF and
 * issue the Bid Capacity Evaluation Sheet (.xlsx): one column per participating
 * bidder, the eleven fixed tender-condition rows, and the 3AN−B bid-capacity
 * block wired as formulas. The header, bidder columns and ECV are filled from
 * the PDF; the per-bidder turnover / N / B / eligibility cells are left blank
 * for the office to key in from each bidder's documents.
 */
export default function EvaluationSheetTab({ office }: Props) {
  const [pdfName, setPdfName] = useState('')
  const [tenderId, setTenderId] = useState('')
  const [nitLine, setNitLine] = useState('')
  const [workLine, setWorkLine] = useState('')
  const [ecv, setEcv] = useState('')
  const [signature, setSignature] = useState('')
  const [bidders, setBidders] = useState<string[]>([])
  const [busy, setBusy] = useState<null | 'pdf' | 'download'>(null)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const pdfInputRef = useRef<HTMLInputElement>(null)

  function defaultSignature(): string {
    const circle = (office.circle ?? '').trim()
    const cno = (office.circleNumber ?? '').trim()
    const corp = (office.corporation ?? '').trim() || 'CMC'
    const circlePart = circle ? `${circle} circle${cno ? `-${cno}` : ''},` : ''
    return `Executive Engineer,\n${circlePart}${corp}`
  }

  async function handlePdf(file: File) {
    setBusy('pdf')
    setError(null)
    setSaved(null)
    try {
      const lines = await pdfToTextLines(file)
      const ev = parseTenderEvaluation(lines)
      const list = parseParticipatingBidders(lines)
      if (list.length === 0) {
        setError(
          `No participating bidders found in ${file.name}. Upload the portal's "View Bidders" (Supplier List) page, saved as PDF.`
        )
        setBusy(null)
        return
      }
      setPdfName(file.name)
      setTenderId(ev.tenderId ?? '')
      setNitLine(
        `NIT No: ${ev.noticeNo ?? ''}${ev.noticeDate ? `,Dt:${ev.noticeDate}` : ''}${ev.tenderId ? `,Tender ID:${ev.tenderId}` : ''}`
      )
      setWorkLine(
        `Name of the work: ${ev.nameOfWork ?? ''}${ev.ecvRupees != null ? `. ecv: ${ev.ecvRupees.toFixed(2)}` : ''}`
      )
      setEcv(ev.ecvRupees != null ? String(ev.ecvRupees) : '')
      setSignature(defaultSignature())
      setBidders(list.map((b) => b.name))
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  function setBidder(i: number, value: string) {
    setBidders((prev) => prev.map((b, j) => (j === i ? value : b)))
  }
  function addBidder() {
    setBidders((prev) => [...prev, ''])
  }
  function removeBidder(i: number) {
    setBidders((prev) => prev.filter((_, j) => j !== i))
  }

  async function download() {
    const names = bidders.map((b) => b.trim()).filter(Boolean)
    if (names.length === 0) {
      setError('Add at least one bidder before downloading.')
      return
    }
    setBusy('download')
    setError(null)
    setSaved(null)
    try {
      const path = await api.exportEvaluationSheet(
        { nitLine, workLine, bidders: names, ecv: Number(ecv.replace(/,/g, '')) || null, signature },
        `Evaluation Sheet${tenderId ? ` - ${tenderId}` : ''}`
      )
      setSaved(path ? `Saved to ${path}` : 'Cancelled.')
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  const ready = bidders.length > 0

  return (
    <div className="card">
      <div className="empty empty--tight">
        <IconChecklist />
        <div className="boq-actions">
          <button className="primary upload-btn" onClick={() => pdfInputRef.current?.click()} disabled={busy === 'pdf'}>
            <IconFolder /> {busy === 'pdf' ? 'Reading PDF…' : pdfName ? 'Change View Bidders PDF' : 'Upload View Bidders PDF'}
          </button>
          <input
            ref={pdfInputRef}
            type="file"
            accept="application/pdf,.pdf"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handlePdf(file)
              e.target.value = ''
            }}
          />
          {ready && (
            <button className="primary" onClick={download} disabled={busy === 'download'}>
              <IconTable /> {busy === 'download' ? 'Building…' : 'Download evaluation sheet'}
            </button>
          )}
        </div>
        {pdfName && <p className="estimate-hint">Bidders read from {pdfName}</p>}
        {error && (
          <div className="notice error">
            <IconChecklist /> {error}
          </div>
        )}
        {saved && (
          <div className="notice ok">
            <IconChecklist /> {saved}
          </div>
        )}
      </div>

      {ready && (
        <div className="estimate-body">
          <div className="loa-manual-fields">
            <label className="field-label">
              NIT / Tender line
              <input value={nitLine} onChange={(e) => setNitLine(e.target.value)} />
            </label>
            <label className="field-label">
              Name of work line
              <input value={workLine} onChange={(e) => setWorkLine(e.target.value)} />
            </label>
            <div className="loa-manual-grid">
              <label className="field-label">
                ECV (rupees)
                <input value={ecv} onChange={(e) => setEcv(e.target.value)} />
              </label>
              <label className="field-label">
                Signature
                <input value={signature.replace(/\n/g, ' — ')} onChange={(e) => setSignature(e.target.value.replace(/ — /g, '\n'))} />
              </label>
            </div>

            <div className="ns-group-title">
              Participating bidders ({bidders.length})
              <button className="ns-add" onClick={addBidder} title="Add a bidder">
                <IconPlus /> Add
              </button>
            </div>
            <div className="ns-bidders">
              {bidders.map((b, i) => (
                <div className="ns-bidder-row" key={i} style={{ gridTemplateColumns: '2rem 1fr 2rem' }}>
                  <span className="mb-serial">{i + 1}</span>
                  <input value={b} onChange={(e) => setBidder(i, e.target.value)} placeholder="Bidder / firm name" />
                  <button className="ns-del" onClick={() => removeBidder(i)} title="Remove">
                    ✕
                  </button>
                </div>
              ))}
            </div>
            <p className="estimate-hint">
              The sheet fills the header, one column per bidder and the ECV, and wires the 3AN−B bid-capacity formula.
              Turnover (A), N, B, the eligibility rows and Remarks are left blank to fill in from each bidder’s documents.
            </p>
          </div>
        </div>
      )}
    </div>
  )
}

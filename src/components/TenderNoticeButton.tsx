import { useMemo, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { IconDoc, IconWarn, IconCheck, IconEye } from './Icons'
import { base64ToUint8 } from './docPage'
import type { ExcelTable } from '@core/types'
import type { BidDocumentBatch } from '@core/types'
import type { TenderNoticeInput } from '../../electron/ipc-contract'

interface Props {
  tables: ExcelTable[]
  /** Called with the NIT number once a notice has been generated and saved. */
  onGenerated?: (nitNo: string) => void
  /** Called once a notice has been generated, with one Bid Document entry per selected Win Code. */
  onBidBatch?: (batch: BidDocumentBatch) => void
}

function findHeader(headers: string[], name: string): string | undefined {
  return headers.find((h) => h.trim().toLowerCase() === name.toLowerCase())
}

/** "Tender notice {leading digits of the NIT No.} of {Circle}", filesystem-sanitized. */
function tenderNoticeFileName(nitNo: string, circle?: string): string {
  const number = nitNo.match(/^\d+/)?.[0] ?? nitNo
  const name = circle ? `Tender notice ${number} of ${circle}` : `Tender notice ${number}`
  return name.replace(/[\\/:*?"<>|]/g, '_').trim()
}

function toDDMMYYYY(iso: string): string {
  const [y, m, d] = iso.split('-')
  return `${d}.${m}.${y}`
}

function todayISO(): string {
  const d = new Date()
  const dd = String(d.getDate()).padStart(2, '0')
  const mm = String(d.getMonth() + 1).padStart(2, '0')
  return `${d.getFullYear()}-${mm}-${dd}`
}

/**
 * "Issue tender notice" — pick one or more Win Codes from the Works List (each
 * becomes a row in the notice's item table, its name/amount pulled straight
 * from the database), a new NIT number, a Dated value (defaults to today,
 * editable), and download/price-bid dates. Preview renders the actual filled
 * document before saving. Times are left at the template's defaults: 2:00
 * P.M download start/end, 2:30 P.M price bid opening (same day as the
 * download end date).
 */
export default function TenderNoticeButton({ tables, onGenerated, onBidBatch }: Props) {
  const worksTable = tables[0] ?? null
  const winCodeHeader = worksTable ? findHeader(worksTable.headers, 'Wincode') : undefined
  const nameHeader = worksTable ? findHeader(worksTable.headers, 'Name of the work') : undefined
  const amountHeader = worksTable ? findHeader(worksTable.headers, 'Amount of estimate') : undefined
  const zoneHeader = worksTable ? findHeader(worksTable.headers, 'Zone') : undefined
  const circleHeader = worksTable ? findHeader(worksTable.headers, 'Circle') : undefined
  const completionHeader = worksTable ? findHeader(worksTable.headers, 'Completion Period') : undefined
  const ecvHeader = worksTable ? findHeader(worksTable.headers, 'ECV') : undefined

  const winCodes = useMemo(() => {
    if (!worksTable || !winCodeHeader) return []
    const seen = new Set<string>()
    for (const row of worksTable.rows) {
      const v = (row[winCodeHeader] ?? '').trim()
      if (v) seen.add(v)
    }
    return Array.from(seen)
  }, [worksTable, winCodeHeader])

  const [open, setOpen] = useState(false)
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [nitNo, setNitNo] = useState('16/DB/EE/Gajularamaram Circle-57/CMC/2026-27')
  const [datedDate, setDatedDate] = useState(todayISO)
  const [startDate, setStartDate] = useState('')
  const [endDate, setEndDate] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const [previewing, setPreviewing] = useState(false)
  const previewRef = useRef<HTMLDivElement>(null)

  const items = useMemo(() => {
    if (!worksTable || !winCodeHeader || !nameHeader || !amountHeader) return []
    return Array.from(selected)
      .map((code) => worksTable.rows.find((r) => (r[winCodeHeader] ?? '').trim() === code))
      .filter((r): r is Record<string, string> => !!r)
      .map((r) => ({ name: r[nameHeader] ?? '', amount: r[amountHeader] ?? '' }))
  }, [worksTable, winCodeHeader, nameHeader, amountHeader, selected])

  // Same selection, same order as `items` above, but carrying the extra
  // fields the Bid Document template needs (ECV/Zone/Circle/Completion
  // Period — EMD @ 1% is computed from ECV, see core/bidDocument.ts) — kept
  // separate from `items` since the tender notice's own item table only
  // ever needs name/amount.
  const bidWorks = useMemo(() => {
    if (!worksTable || !winCodeHeader || !nameHeader || !amountHeader) return []
    return Array.from(selected)
      .map((code) => worksTable.rows.find((r) => (r[winCodeHeader] ?? '').trim() === code))
      .filter((r): r is Record<string, string> => !!r)
      .map((r, i) => ({
        serial: i + 1,
        winCode: winCodeHeader ? (r[winCodeHeader] ?? '').trim() : undefined,
        name: r[nameHeader] ?? '',
        amount: r[amountHeader] ?? '',
        ecv: ecvHeader ? r[ecvHeader] : undefined,
        zone: zoneHeader ? r[zoneHeader] : undefined,
        circle: circleHeader ? r[circleHeader] : undefined,
        completionPeriod: completionHeader ? r[completionHeader] : undefined
      }))
  }, [
    worksTable,
    winCodeHeader,
    nameHeader,
    amountHeader,
    ecvHeader,
    zoneHeader,
    circleHeader,
    completionHeader,
    selected
  ])

  function toggleWinCode(code: string) {
    setSelected((prev) => {
      const next = new Set(prev)
      if (next.has(code)) next.delete(code)
      else next.add(code)
      return next
    })
  }

  function openModal() {
    setOpen(true)
    setError(null)
    setSaved(null)
    setPreviewing(false)
  }

  function buildInput(): TenderNoticeInput | null {
    if (selected.size === 0 || !nitNo.trim() || !datedDate || !startDate || !endDate) {
      setError('Select at least one Win Code, and fill in the NIT number, Dated, and both dates.')
      return null
    }
    return {
      nitNo: nitNo.trim(),
      items,
      startDate: toDDMMYYYY(startDate),
      endDate: toDDMMYYYY(endDate),
      today: toDDMMYYYY(datedDate)
    }
  }

  async function preview() {
    setError(null)
    const input = buildInput()
    if (!input) return
    setBusy(true)
    try {
      const b64 = await api.previewTenderNotice(input)
      setPreviewing(true)
      // Wait a tick for the preview container to mount.
      requestAnimationFrame(() => {
        void (async () => {
          const container = previewRef.current
          if (!container) return
          container.innerHTML = ''
          await renderAsync(base64ToUint8(b64), container, undefined, {
            className: 'docx',
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false,
            breakPages: true,
            experimental: true,
            trimXmlDeclaration: true,
            useBase64URL: true,
            renderHeaders: true,
            renderFooters: true,
            renderFootnotes: true,
            renderEndnotes: true,
            renderChanges: false
          })
        })()
      })
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  async function generate() {
    setError(null)
    const input = buildInput()
    if (!input) return
    setBusy(true)
    try {
      const suggestedName = tenderNoticeFileName(input.nitNo, bidWorks[0]?.circle)
      const path = await api.generateTenderNotice(input, suggestedName)
      if (path) {
        setSaved(path)
        setOpen(false)
        onGenerated?.(input.nitNo)
        onBidBatch?.({
          id: `bid-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
          nitNo: input.nitNo,
          dated: input.today,
          downloadStartDate: input.startDate,
          downloadEndDate: input.endDate,
          works: bidWorks,
          createdDate: todayISO()
        })
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <button className="primary" onClick={openModal}>
        <IconDoc /> Issue tender notice
      </button>

      {saved && !open && (
        <div className="notice ok">
          <IconCheck />
          Saved to {saved}
        </div>
      )}

      {open && (
        <div className="editor-overlay" onClick={() => setOpen(false)}>
          <div
            className={`tender-modal ${previewing ? 'tender-modal-wide' : ''}`}
            onClick={(e) => e.stopPropagation()}
          >
            {previewing ? (
              <>
                <h3>Preview — this is exactly what gets saved</h3>
                <div className="tender-preview-scroll">
                  <div ref={previewRef} className="docprev-body" />
                </div>
                {error && (
                  <div className="notice error">
                    <IconWarn />
                    {error}
                  </div>
                )}
                <div className="confirm-actions">
                  <button className="ghost" onClick={() => setPreviewing(false)} disabled={busy}>
                    Back to edit
                  </button>
                  <button className="primary" onClick={generate} disabled={busy}>
                    {busy ? 'Generating…' : 'Generate'}
                  </button>
                </div>
              </>
            ) : (
              <>
                <h3>Issue tender notice</h3>

                {winCodes.length > 0 ? (
                  <div className="tender-field">
                    <span>Win Code(s)</span>
                    <div className="tender-wincode-list">
                      {winCodes.map((code) => (
                        <label key={code} className="tender-wincode-item">
                          <input
                            type="checkbox"
                            checked={selected.has(code)}
                            onChange={() => toggleWinCode(code)}
                          />
                          {code}
                        </label>
                      ))}
                    </div>
                  </div>
                ) : (
                  <p className="hint">No Win Codes found — add works to the Works List first.</p>
                )}

                <div className="tender-field-row">
                  <label className="tender-field tender-field-nitno">
                    <span>NIT No.</span>
                    <input
                      value={nitNo}
                      onChange={(e) => setNitNo(e.target.value)}
                      placeholder="16/DB/EE/Gajularamaram Circle-57/CMC/2026-27"
                    />
                  </label>

                  <label className="tender-field tender-field-dated">
                    <span>Dated</span>
                    <input type="date" value={datedDate} onChange={(e) => setDatedDate(e.target.value)} />
                  </label>
                </div>

                <label className="tender-field">
                  <span>Download Start Date (2:00 P.M)</span>
                  <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
                </label>

                <label className="tender-field">
                  <span>Download End Date (2:00 P.M)</span>
                  <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
                </label>

                <p className="hint">
                  Price Bid Opening is the same day as the download end date, at 2:30 P.M.
                </p>

                {error && (
                  <div className="notice error">
                    <IconWarn />
                    {error}
                  </div>
                )}

                <div className="confirm-actions">
                  <button className="ghost" onClick={() => setOpen(false)} disabled={busy}>
                    Cancel
                  </button>
                  <button className="ghost" onClick={preview} disabled={busy}>
                    <IconEye /> {busy ? 'Loading…' : 'Preview'}
                  </button>
                  <button className="primary" onClick={generate} disabled={busy}>
                    {busy ? 'Generating…' : 'Generate'}
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </>
  )
}

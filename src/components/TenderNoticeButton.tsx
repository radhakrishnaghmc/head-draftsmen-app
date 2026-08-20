import { useEffect, useMemo, useRef, useState } from 'react'
import { api } from '../ipc'
import { IconDoc, IconWarn, IconCheck, IconEye } from './Icons'
import { base64ToUint8, renderDocPreview } from './docPage'
import type { ExcelTable } from '@core/types'
import type { BidDocumentBatch } from '@core/types'
import type { TenderNoticeInput } from '../../electron/ipc-contract'
import { indianFinancialYear } from '@core/workOrderAgreement'
import { type Office, officeScopedKey, isZoneOnlyOffice, CONTACT_KEYS } from '../office'
import { closeOnBackdropMouseDown } from '../overlayClose'

interface Props {
  tables: ExcelTable[]
  /** The chosen office (Works List page) — the default NIT No is built from its Circle/CNO/Corporation. */
  office?: Office
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

/** The office's default e-mail, following the circle's number ("eec58.GHMC@gmail.com"). */
function defaultEmailFor(office?: Office): string {
  return office?.circleNumber ? `eec${office.circleNumber}.GHMC@gmail.com` : ''
}

/**
 * Whether the full NIT No can be composed from the office — needs its
 * Corporation, plus either (Circle + Circle number) or, for a Zone-level (SE)
 * office with no circle of its own, just the Zone. When it can, the user only
 * supplies the running serial and the rest comes from the office, so the NIT
 * number can never disagree with the rest of the document.
 */
function canComposeNit(office?: Office): boolean {
  if (!office?.corporation) return false
  if (office.circle && office.circleNumber) return true
  return !!(office.zone && !office.circle)
}

/** yyyy-mm-dd (an <input type="date">'s raw value) -> a local-time Date, so the financial year comes out right regardless of timezone (new Date("yyyy-mm-dd") parses as UTC midnight, which can roll to the wrong local day). Blank/invalid -> today. */
function parseIsoDateLocal(iso: string): Date {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso)
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date()
}

/**
 * Build the full NIT No from the running serial + the office's
 * circle-or-zone/corporation/financial-year. The financial year comes from
 * the notice's own Dated field (falling back to today when it's blank) —
 * not today's date — so it's the NIT's actual financial year, and changing
 * Dated updates it live.
 */
function composeNitNo(serial: string, office: Office | undefined, datedIso: string): string {
  const s = serial.trim() || '16'
  // A Zone-level office has no Circle of its own — name the Zone instead of
  // "Circle Circle-<CNO>", which would otherwise print "undefined".
  const place = office?.circle ? `${office.circle} Circle-${office.circleNumber}` : `${office?.zone} Zone`
  return `${s}/DB/EE/${place}/${office?.corporation}/${indianFinancialYear(parseIsoDateLocal(datedIso))}`
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
export default function TenderNoticeButton({ tables, office, onGenerated, onBidBatch }: Props) {
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

  // The NIT No default follows the chosen office — its Circle, Circle number
  // and Corporation — the circle, corporation and financial year are pulled from
  // the office (the single source of truth for circle/zone across the whole
  // notice); the user only supplies the running serial ("16"). When the office
  // isn't fully chosen we fall back to a hand-typed full NIT No.
  const composeReady = canComposeNit(office)

  // Running serial only (office-composed path); the full hand-typed NIT No
  // (fallback path when the office isn't fully chosen).
  const [nitSerial, setNitSerial] = useState('16')
  // Starts blank — only an example ever showed in the `placeholder` below —
  // so a user who doesn't type their own NIT No can't accidentally submit
  // this example text as if it were real.
  const [nitNoManual, setNitNoManual] = useState('')
  const [datedDate, setDatedDate] = useState(todayISO)
  // Not memoized — the financial year depends on `datedDate` (live) and,
  // when that's blank, on today's real date, so it needs to re-run on every
  // render rather than only when nitSerial/office/datedDate change.
  const nitNo = composeReady ? composeNitNo(nitSerial, office, datedDate) : nitNoManual

  const [startDate, setStartDate] = useState('')
  const [endDate, setEndDate] = useState('')
  const [email, setEmail] = useState('')
  const [eePhone, setEePhone] = useState('')
  const [hdPhone, setHdPhone] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const [previewing, setPreviewing] = useState(false)
  const previewRef = useRef<HTMLDivElement>(null)
  const winListRef = useRef<HTMLDivElement>(null)

  // New works are appended to the Works List, so their Win Codes land at the end
  // of this list — scroll to the bottom when the dialog opens so the latest ones
  // are in view by default.
  useEffect(() => {
    if (open && !previewing && winListRef.current) {
      winListRef.current.scrollTop = winListRef.current.scrollHeight
    }
  }, [open, previewing, winCodes.length])

  // Reset the NIT No to the current office's default each time the dialog opens,
  // so it always reflects the chosen Circle/Corporation. Contact details are
  // pre-filled from the last-remembered values (or, for the e-mail, derived from
  // the circle number when nothing's been saved yet).
  useEffect(() => {
    if (!open) return
    setNitSerial('16')
    setEmail(localStorage.getItem(officeScopedKey(CONTACT_KEYS.email, office)) || defaultEmailFor(office))
    setEePhone(localStorage.getItem(officeScopedKey(CONTACT_KEYS.eePhone, office)) || '')
    setHdPhone(localStorage.getItem(officeScopedKey(CONTACT_KEYS.hdPhone, office)) || '')
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open])

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
      today: toDDMMYYYY(datedDate),
      // Carry the chosen office's circle/zone so the template's hard-coded
      // Gajularamaram Circle-57 / Quthbullapur Zone place names are rewritten
      // to the circle this notice is actually being issued for.
      circle: office?.circle,
      circleNumber: office?.circleNumber,
      zone: office?.zone,
      email: email.trim(),
      eePhone: eePhone.trim(),
      hdPhone: hdPhone.trim()
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
          await renderDocPreview(base64ToUint8(b64), container)
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
        // Remember the contact details for next time (per machine, per office).
        localStorage.setItem(officeScopedKey(CONTACT_KEYS.email, office), email.trim())
        localStorage.setItem(officeScopedKey(CONTACT_KEYS.eePhone, office), eePhone.trim())
        localStorage.setItem(officeScopedKey(CONTACT_KEYS.hdPhone, office), hdPhone.trim())
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
        <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => setOpen(false))}>
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
                    {busy ? 'Issuing…' : 'Issue'}
                  </button>
                </div>
              </>
            ) : (
              <>
                <h3>Issue tender notice</h3>

                {winCodes.length > 0 ? (
                  <div className="tender-field">
                    <span>Win Code(s)</span>
                    <div className="tender-wincode-list" ref={winListRef}>
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
                  {composeReady ? (
                    <label className="tender-field tender-field-nitno">
                      <span>NIT Serial No.</span>
                      <input
                        value={nitSerial}
                        onChange={(e) => setNitSerial(e.target.value)}
                        placeholder="16"
                      />
                    </label>
                  ) : (
                    <label className="tender-field tender-field-nitno">
                      <span>NIT No.</span>
                      <input
                        value={nitNoManual}
                        onChange={(e) => setNitNoManual(e.target.value)}
                        placeholder="16/DB/EE/Gajularamaram Circle-57/CMC/2026-27"
                      />
                    </label>
                  )}

                  <label className="tender-field tender-field-dated">
                    <span>Dated</span>
                    <input type="date" value={datedDate} onChange={(e) => setDatedDate(e.target.value)} />
                  </label>
                </div>

                {composeReady && (
                  <p className="hint">
                    NIT No.: <strong>{nitNo}</strong> — circle & corporation come from the office
                    selected on the Works List.
                  </p>
                )}

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

                <label className="tender-field">
                  <span>Email ID</span>
                  <input
                    type="email"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    placeholder="eec58.GHMC@gmail.com"
                  />
                </label>

                <div className="tender-field-row">
                  <label className="tender-field">
                    <span>{isZoneOnlyOffice(office) ? 'Superintending Engineer Phone' : 'Executive Engineer Phone'}</span>
                    <input
                      type="tel"
                      value={eePhone}
                      onChange={(e) => setEePhone(e.target.value)}
                      placeholder="10-digit mobile number"
                    />
                  </label>

                  <label className="tender-field">
                    <span>Head Draughtsman Phone</span>
                    <input
                      type="tel"
                      value={hdPhone}
                      onChange={(e) => setHdPhone(e.target.value)}
                      placeholder="10-digit mobile number"
                    />
                  </label>
                </div>

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
                    {busy ? 'Issuing…' : 'Issue'}
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

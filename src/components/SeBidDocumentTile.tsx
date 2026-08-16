import { useMemo, useRef, useState } from 'react'
import { renderAsync } from '../lazyDocxPreview'
import { api } from '../ipc'
import { IconDoc, IconEye, IconDownload, IconWarn, IconCheck } from './Icons'
import { base64ToUint8 } from './docPage'
import type { ExcelTable } from '@core/types'
import type { BidDocumentInput } from '../../electron/ipc-contract'
import { type Office } from '../office'
import { closeOnBackdropMouseDown } from '../overlayClose'

interface Props {
  tables: ExcelTable[]
  office?: Office
}

function findHeader(headers: string[], name: string): string | undefined {
  return headers.find((h) => h.trim().toLowerCase() === name.toLowerCase())
}

function todayISO(): string {
  const d = new Date()
  const mm = String(d.getMonth() + 1).padStart(2, '0')
  const dd = String(d.getDate()).padStart(2, '0')
  return `${d.getFullYear()}-${mm}-${dd}`
}

function toDDMMYYYY(iso: string): string {
  const [y, m, d] = iso.split('-')
  return `${d}.${m}.${y}`
}

interface FormState {
  nitNo: string
  dated: string
  amount: string
  ecv: string
  completionPeriod: string
  itemNo: string
  tsNo: string
  tsDate: string
  asAuthority: 'zonal' | 'commissioner'
  asDate: string
}

const EMPTY_FORM: FormState = {
  nitNo: '',
  dated: todayISO(),
  amount: '',
  ecv: '',
  completionPeriod: '',
  itemNo: '',
  tsNo: '',
  tsDate: '',
  asAuthority: 'commissioner',
  asDate: ''
}

/**
 * Standalone "Bid Document" entry point for an SE (zone-only, no Circle)
 * office — shown below the Calendar workspace so an SE office can generate
 * its Bid Document straight from a Works List row, without first issuing a
 * tender notice (unlike the EE flow, which only offers Bid Documents as a
 * byproduct of TenderNoticeButton's batch). Picking a work pre-fills
 * Amount/ECV/Completion Period from its Works List row when present; ECV
 * stays editable so it can be typed in when the row doesn't have one yet.
 */
export default function SeBidDocumentTile({ tables, office }: Props) {
  const seMode = !!office?.zone?.trim() && !office?.circle?.trim()
  const worksTable = tables[0] ?? null
  const nameHeader = worksTable ? findHeader(worksTable.headers, 'Name of the work') : undefined
  const amountHeader = worksTable ? findHeader(worksTable.headers, 'Amount of estimate') : undefined
  const ecvHeader = worksTable ? findHeader(worksTable.headers, 'ECV') : undefined
  const zoneHeader = worksTable ? findHeader(worksTable.headers, 'Zone') : undefined
  const completionHeader = worksTable ? findHeader(worksTable.headers, 'Completion Period') : undefined

  const workNames = useMemo(() => {
    if (!worksTable || !nameHeader) return []
    const seen = new Set<string>()
    for (const row of worksTable.rows) {
      const v = (row[nameHeader] ?? '').trim()
      if (v) seen.add(v)
    }
    return Array.from(seen)
  }, [worksTable, nameHeader])

  const [open, setOpen] = useState(false)
  const [workName, setWorkName] = useState('')
  const [zone, setZone] = useState('')
  const [form, setForm] = useState<FormState>(EMPTY_FORM)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saved, setSaved] = useState<string | null>(null)
  const [previewing, setPreviewing] = useState(false)
  const previewRef = useRef<HTMLDivElement>(null)

  if (!seMode) return null

  function set<K extends keyof FormState>(key: K, value: FormState[K]) {
    setForm((prev) => ({ ...prev, [key]: value }))
  }

  function openModal() {
    setOpen(true)
    setError(null)
    setSaved(null)
    setPreviewing(false)
    setWorkName('')
    setZone(office?.zone ?? '')
    setForm(EMPTY_FORM)
  }

  // Pre-fill Amount/ECV/Zone/Completion Period from the matching Works List
  // row — every field stays editable afterwards, so a missing ECV (or any
  // other field) is simply typed in rather than gated behind a separate step.
  function pickWork(name: string) {
    setWorkName(name)
    if (!worksTable || !nameHeader) return
    const row = worksTable.rows.find((r) => (r[nameHeader] ?? '').trim() === name)
    if (!row) return
    setForm((prev) => ({
      ...prev,
      amount: amountHeader ? (row[amountHeader] ?? '') : prev.amount,
      ecv: ecvHeader ? (row[ecvHeader] ?? '') : prev.ecv,
      completionPeriod: completionHeader ? (row[completionHeader] ?? '') : prev.completionPeriod
    }))
    setZone(zoneHeader ? (row[zoneHeader] ?? '') || (office?.zone ?? '') : office?.zone ?? '')
  }

  function buildInput(): BidDocumentInput | null {
    if (!workName.trim() || !form.nitNo.trim() || !form.dated) {
      setError('Pick a work, and fill in the NIT number and Dated.')
      return null
    }
    return {
      nitNo: form.nitNo.trim(),
      dated: toDDMMYYYY(form.dated),
      downloadStartDate: '',
      downloadEndDate: '',
      work: {
        serial: 1,
        name: workName,
        amount: form.amount,
        ecv: form.ecv,
        zone,
        completionPeriod: form.completionPeriod,
        itemNo: form.itemNo,
        tsNo: form.tsNo,
        tsDate: form.tsDate,
        asAuthority: form.asAuthority,
        asDate: form.asDate
      }
    }
  }

  async function preview() {
    setError(null)
    const input = buildInput()
    if (!input) return
    setBusy(true)
    try {
      const b64 = await api.previewBidDocument(input)
      setPreviewing(true)
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

  async function save() {
    setError(null)
    const input = buildInput()
    if (!input) return
    setBusy(true)
    try {
      const suggestedName = `Bid Document ${workName}`.slice(0, 150)
      const path = await api.generateBidDocument(input, suggestedName)
      if (path) {
        setSaved(path)
        setOpen(false)
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <section className="card bid-batch">
        <div className="card-head">
          <div className="head-ic">
            <IconDoc />
          </div>
          <div className="titles">
            <h2>Bid Document</h2>
            <p className="sub">SE office — generate directly from a Works List entry.</p>
          </div>
          <button className="primary" onClick={openModal}>
            <IconDoc /> Issue Bid Document
          </button>
        </div>

        {saved && !open && (
          <div className="notice ok">
            <IconCheck />
            Saved to {saved}
          </div>
        )}
      </section>

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
                  <button className="primary" onClick={save} disabled={busy}>
                    {busy ? 'Saving…' : 'Save'}
                  </button>
                </div>
              </>
            ) : (
              <>
                <h3>Issue SE Bid Document</h3>

                <label className="tender-field">
                  <span>Name of the work</span>
                  {workNames.length > 0 ? (
                    <select value={workName} onChange={(e) => pickWork(e.target.value)}>
                      <option value="">Select a work…</option>
                      {workNames.map((n) => (
                        <option key={n} value={n}>
                          {n}
                        </option>
                      ))}
                    </select>
                  ) : (
                    <input value={workName} onChange={(e) => setWorkName(e.target.value)} placeholder="Name of the work" />
                  )}
                </label>

                <div className="tender-field-row">
                  <label className="tender-field tender-field-nitno">
                    <span>NIT No.</span>
                    <input
                      value={form.nitNo}
                      onChange={(e) => set('nitNo', e.target.value)}
                      placeholder="13/SE/QBZ/CMC/2026-27"
                    />
                  </label>
                  <label className="tender-field tender-field-dated">
                    <span>Dated</span>
                    <input type="date" value={form.dated} onChange={(e) => set('dated', e.target.value)} />
                  </label>
                </div>

                <div className="tender-field-row bid-se-fields">
                  <label className="tender-field">
                    <span>Amount of Estimate (Lakhs)</span>
                    <input value={form.amount} onChange={(e) => set('amount', e.target.value)} placeholder="5" />
                  </label>
                  <label className="tender-field">
                    <span>ECV (Rs.)</span>
                    <input value={form.ecv} onChange={(e) => set('ecv', e.target.value)} placeholder="415020" />
                  </label>
                </div>

                <div className="tender-field-row bid-se-fields">
                  <label className="tender-field">
                    <span>Item No.</span>
                    <input value={form.itemNo} onChange={(e) => set('itemNo', e.target.value)} placeholder="7" />
                  </label>
                  <label className="tender-field">
                    <span>Completion Period</span>
                    <input
                      value={form.completionPeriod}
                      onChange={(e) => set('completionPeriod', e.target.value)}
                      placeholder="2"
                    />
                  </label>
                </div>

                <label className="tender-field">
                  <span>T.S. No.</span>
                  <input
                    value={form.tsNo}
                    onChange={(e) => set('tsNo', e.target.value)}
                    placeholder="29/SE/QBZ/CMC/2026-27"
                  />
                </label>

                <div className="tender-field-row bid-se-fields">
                  <label className="tender-field">
                    <span>T.S. Date</span>
                    <input value={form.tsDate} onChange={(e) => set('tsDate', e.target.value)} placeholder="04-07-2026" />
                  </label>
                  <label className="tender-field">
                    <span>AS Authority</span>
                    <select value={form.asAuthority} onChange={(e) => set('asAuthority', e.target.value as 'zonal' | 'commissioner')}>
                      <option value="commissioner">Commissioner, CMC</option>
                      <option value="zonal">Zonal Commissioner</option>
                    </select>
                  </label>
                </div>

                <label className="tender-field">
                  <span>AS Date</span>
                  <input value={form.asDate} onChange={(e) => set('asDate', e.target.value)} placeholder="23.06.2026" />
                </label>

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
                  <button className="primary" onClick={save} disabled={busy}>
                    {busy ? 'Saving…' : 'Save'}
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

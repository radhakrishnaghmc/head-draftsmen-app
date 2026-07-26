import { useEffect, useMemo, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import { api } from '../ipc'
import { parseIntimationNotice, type IntimationNotice } from '@core/intimationNotice'
import { computeWorkAmounts, indianDigitGroups } from '@core/worksAmounts'
import type { PlaceholderMatch } from '@core/createDocument'
import { base64ToUint8, DOCX_PREVIEW_OPTIONS, PAGE_WIDTH } from './docPage'
import { IconFolder, IconDownload, IconPrint, IconWarn, IconBell } from './Icons'
import type { ExcelTable } from '@core/types'

interface Props {
  tables: ExcelTable[]
}

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

function rowLabel(row: Record<string, string>, headers: string[], index: number): string {
  const nameCol = headers.find((h) => /name of (the )?work/i.test(h)) ?? headers[0]
  const v = nameCol ? (row[nameCol] ?? '').trim() : ''
  return v ? `${index + 1}. ${v}` : `Row ${index + 1}`
}

function group(n: number | null | undefined): string {
  return n == null ? '' : indianDigitGroups(Math.round(n))
}

/**
 * Resolves one Intimation placeholder's value: the uploaded portal notice
 * (see core/intimationNotice.ts) wins for what it carries — agency name,
 * address, NIT No, ECV, contract value — with the picked Works List row
 * filling the rest (Circle, phone, EMD, ASD, Tender %). ECV/Contract/EMD/ASD
 * come back as bare Indian-grouped figures, since the template already
 * prints the "Rs …/-" wording around each placeholder. Every value is
 * pre-filled but editable, so an unmatched or misread field can be fixed
 * before the letter is generated.
 */
function resolveValue(
  label: string,
  notice: IntimationNotice,
  row: Record<string, string>
): string {
  const c = computeWorkAmounts(row)
  switch (norm(label)) {
    case 'agency name':
    case 'name of the agency':
      return notice.agencyName ?? row['Name of the Agency'] ?? ''
    case 'address of the agency':
      return notice.address ?? row['Address of the agency'] ?? ''
    case 'agency phone number':
    case 'phone number of the agency':
      return row['Phone number of the agency'] ?? ''
    case 'circle':
      return row['Circle'] ?? ''
    case 'cno':
      return row['CNO'] ?? ''
    case 'zone':
      return row['Zone'] ?? ''
    case 'name of the work':
      return row['Name of the work'] ?? ''
    case 'nit no':
    case 'tender notice no':
      return notice.nitNo ?? row['Tender Notice No'] ?? ''
    case 'tender pencentage':
    case 'tender percentage':
      return (row['Tender Percentage'] ?? '').replace(/%/g, '').trim()
    case 'ecv':
      return notice.ecvRupees != null ? group(notice.ecvRupees) : group(c.ecv)
    case 'contract amount':
      return notice.contractRupees != null ? group(notice.contractRupees) : group(c.contractAmount)
    case 'emd':
    case 'emd 1.5%':
      return group(c.emd1_5)
    case 'emd 1%':
      return group(c.emd1)
    case 'asd':
      return group(c.asd)
    default:
      return ''
  }
}

/**
 * Give Intimation — fills the bundled Intimation format (a .docx mail-merge
 * template with {{placeholders}}) for a Works List work: most fields come
 * from the picked Works List row, while the agency's address (and, when
 * present, agency name / NIT No / ECV / contract value) come from an uploaded
 * portal "View Intimation Notice" page (.html). Reuses the same docx
 * placeholder-fill + docx-preview + export/print pipeline as Issue Document.
 */
export default function GiveIntimationTab({ tables }: Props) {
  const table = tables[0] ?? null

  const [templateB64, setTemplateB64] = useState<string | null>(null)
  const [labels, setLabels] = useState<string[]>([])
  const [loadError, setLoadError] = useState<string | null>(null)

  const [rowIndex, setRowIndex] = useState(0)
  const [notice, setNotice] = useState<IntimationNotice | null>(null)
  const [noticeName, setNoticeName] = useState('')
  const [values, setValues] = useState<Record<string, string>>({})
  // Labels the user has hand-edited — kept out of the auto-fill so a re-pick
  // or re-upload never clobbers a manual correction.
  const [touched, setTouched] = useState<Record<string, boolean>>({})

  const [busy, setBusy] = useState<null | 'preview' | 'download' | 'print'>(null)
  const [actionError, setActionError] = useState<string | null>(null)
  const [actionSaved, setActionSaved] = useState<string | null>(null)
  const [previewPages, setPreviewPages] = useState(0)

  const fileInputRef = useRef<HTMLInputElement>(null)
  const previewRef = useRef<HTMLDivElement>(null)
  const printScratchRef = useRef<HTMLDivElement>(null)

  const selectedRow = table && table.rows.length > 0 ? table.rows[Math.min(rowIndex, table.rows.length - 1)] : null

  // Load the bundled Intimation format once, and read its placeholders.
  useEffect(() => {
    let cancelled = false
    void (async () => {
      try {
        const b64 = await api.intimationTemplate()
        const found = await api.findPlaceholdersInDocument(b64)
        if (cancelled) return
        setTemplateB64(b64)
        setLabels(found)
      } catch (e) {
        if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  // Re-fill every untouched placeholder whenever the picked row or uploaded
  // notice changes — a hand-edited field (touched) is left exactly as typed.
  useEffect(() => {
    if (labels.length === 0) return
    setValues((prev) => {
      const next = { ...prev }
      for (const label of labels) {
        if (touched[label]) continue
        next[label] = resolveValue(label, notice ?? {}, selectedRow ?? {})
      }
      return next
    })
    // selectedRow identity changes with rowIndex/table; notice with each upload.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [labels, rowIndex, notice, table])

  function updateValue(label: string, value: string) {
    setTouched((prev) => ({ ...prev, [label]: true }))
    setValues((prev) => ({ ...prev, [label]: value }))
  }

  async function handleNoticeFile(file: File) {
    setActionError(null)
    try {
      const text = await file.text()
      setNotice(parseIntimationNotice(text))
      setNoticeName(file.name)
      setTouched({}) // a fresh notice supersedes prior manual edits
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    }
  }

  async function fillTemplate(): Promise<string> {
    if (!templateB64) throw new Error('Intimation format not loaded yet.')
    // Fill each {{Label}} directly from its edited value: map every label to
    // itself so fillPlaceholdersInDocument reads values[label] (rather than
    // resolving through a Works List column).
    const resolved: PlaceholderMatch[] = labels.map((label) => ({ label, column: label, score: 1 }))
    return api.fillPlaceholdersInDocument(templateB64, resolved, values)
  }

  async function showPreview() {
    setBusy('preview')
    setActionError(null)
    try {
      const filled = await fillTemplate()
      const container = previewRef.current
      if (!container) return
      container.innerHTML = ''
      await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
      setPreviewPages(container.querySelectorAll('section.docx').length)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function downloadIntimation(formats: ('docx' | 'pdf')[]) {
    setBusy('download')
    setActionError(null)
    setActionSaved(null)
    try {
      const filled = await fillTemplate()
      const name = `Intimation${selectedRow?.['Name of the Agency'] ? ` - ${selectedRow['Name of the Agency']}` : ''}`
      const res = await api.exportCreatedDocument(filled, name, formats)
      setActionSaved(res && res.length > 0 ? `Saved: ${res.map((r) => r.file).join(', ')}` : 'Cancelled.')
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  async function printIntimation() {
    setBusy('print')
    setActionError(null)
    try {
      const filled = await fillTemplate()
      const container = printScratchRef.current
      if (!container) throw new Error('Print failed to initialize.')
      container.innerHTML = ''
      await renderAsync(base64ToUint8(filled), container, undefined, DOCX_PREVIEW_OPTIONS)
      await api.printCreatedDocument(container.innerHTML)
    } catch (e) {
      setActionError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(null)
    }
  }

  // Auto-refresh the preview whenever the filled values change.
  useEffect(() => {
    if (!templateB64 || labels.length === 0) return
    void showPreview()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [templateB64, values])

  const noWorks = !table || table.rows.length === 0

  return (
    <div className="card">
      <div ref={printScratchRef} style={{ position: 'fixed', top: -99999, left: -99999, width: PAGE_WIDTH }} aria-hidden />

      <div className="empty">
        <IconBell />
        <p>
          Pick the work below and upload the portal's <strong>Intimation Notice</strong> (.html) — the app fills the
          Intimation letter, taking the agency address from the notice and the rest from the Works List.
        </p>
        <div className="boq-actions">
          <button className="primary" onClick={() => fileInputRef.current?.click()} disabled={!templateB64}>
            <IconFolder /> {notice ? 'Change Intimation Notice (.html)' : 'Upload Intimation Notice (.html)'}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".html,.htm,text/html"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file) void handleNoticeFile(file)
              e.target.value = ''
            }}
          />
        </div>
        {noticeName && <p className="estimate-hint">Address read from {noticeName}</p>}
      </div>

      {loadError && (
        <div className="notice error">
          <IconWarn /> {loadError}
        </div>
      )}

      {noWorks ? (
        <div className="notice">Add works to the Works List first — the Intimation letter is filled from a work's row.</div>
      ) : (
        <div className="estimate-body">
          <div className="estimate-preview">
            <span className="estimate-preview-title">Live Preview{previewPages > 1 ? ` — ${previewPages} pages` : ''}</span>
            <div className="estimate-preview-scroll">
              <div ref={previewRef} className="intimation-docx-preview" />
            </div>
            <div className="doc-sheet-footer">
              <span className="estimate-hint">Updates live as you fill in the details.</span>
              <button className="primary" onClick={() => downloadIntimation(['docx'])} disabled={busy !== null}>
                <IconDownload /> {busy === 'download' ? 'Saving…' : 'Word'}
              </button>
              <button className="primary" onClick={() => downloadIntimation(['pdf'])} disabled={busy !== null}>
                <IconDownload /> {busy === 'download' ? 'Saving…' : 'PDF'}
              </button>
              <button className="primary" onClick={printIntimation} disabled={busy !== null}>
                <IconPrint /> {busy === 'print' ? 'Opening…' : 'Print'}
              </button>
            </div>
            {actionError && (
              <div className="notice error">
                <IconWarn /> {actionError}
              </div>
            )}
            {actionSaved && <p className="estimate-hint">{actionSaved}</p>}
          </div>

          <div className="estimate-details">
            <label className="estimate-details-field">
              Work
              <select className="editor-name" value={rowIndex} onChange={(e) => setRowIndex(Number(e.target.value))}>
                {table!.rows.map((row, i) => (
                  <option value={i} key={i}>
                    {rowLabel(row, table!.headers, i)}
                  </option>
                ))}
              </select>
            </label>

            <span className="estimate-preview-title">Details filled in</span>
            {labels.length === 0 ? (
              <p className="estimate-hint">No {'{{placeholders}}'} found in the Intimation format.</p>
            ) : (
              labels.map((label) => (
                <label className="estimate-details-field" key={label}>
                  {label}
                  {norm(label) === 'address of the agency' && (
                    <span className="estimate-hint">— from the uploaded notice</span>
                  )}
                  <input
                    className="editor-name"
                    placeholder={label}
                    value={values[label] ?? ''}
                    onChange={(e) => updateValue(label, e.target.value)}
                  />
                </label>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  )
}

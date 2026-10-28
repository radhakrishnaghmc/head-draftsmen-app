import { useEffect, useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import { closeOnBackdropMouseDown } from '../overlayClose'
import { type Office, withZoneCircles } from '../office'
import type { SheetGrid } from '@core/sheet'
import type { MfStatusKey, MonitoringFormatSummary, MonitoringFormatWorkRow } from '@core/monitoringFormat'
import {
  extractMonitoringFormatForOffice,
  extractMonitoringFormatWorksForOffice,
  monitoringFormatWorksTable
} from '@core/monitoringFormat'
import { findMonitoringFormatErrors } from '@core/worksListAgent'
import { indianDigitGroups } from '@core/worksAmounts'
import { api } from '../ipc'
import { IconPlus, IconLink, IconWarn, IconCheck, IconClose } from './Icons'

interface Props {
  office: Office
  monitoringFormat?: MonitoringFormatSummary
  monitoringFormatWorks?: MonitoringFormatWorkRow[]
  onImportMonitoringFormat: (summary: MonitoringFormatSummary, works: MonitoringFormatWorkRow[]) => void
  monitoringFormatLink?: string
  onSaveMonitoringFormatLink: (url: string) => void
  /** After a successful import from a link, with the workbook's sheets: lets the app copy the link + each circle's own sheet to the circles under a zone-level (SE) office. Returns a short note to show, if any. */
  /** Deletes this office's imported Monitoring Format data (the app asks for confirmation). */
  onDeleteMonitoringFormat?: () => void
  /** Headings of the works-table columns the user deleted; they are left out of the table and of every download. */
  hiddenColumns?: string[]
  onHiddenColumnsChange?: (headers: string[]) => void
  onImportedFromLink?: (sheets: SheetGrid[], url: string) => string | void
  /** Click a status tile: save a landscape PDF of just that status's works ('all' = every work). Resolves true once saved. */
  onExportStatusPdf?: (key: MfStatusKey | 'all', label: string) => Promise<boolean>
  /** Download one PDF of the works in ANY combination of statuses the user ticked. Resolves true once saved. */
  onExportCombinedPdf?: (keys: MfStatusKey[], label: string) => Promise<boolean>
  /** Click the Abstract tile: save the sheet's Abstract (item type + ward/circle wise) as a PDF. Resolves true once saved. */
  onExportAbstractPdf?: () => Promise<boolean>
  /** Download exactly the works the table's column filters leave. Resolves true once saved. */
  onExportFilteredWorks?: (works: MonitoringFormatWorkRow[], format: 'xlsx' | 'pdf') => Promise<boolean>
  /** The works rows are being re-read from the saved link in the background. */
  worksLoading?: boolean
}

/** Sorts the verify messages under the part of the sheet they are about, in the order the user reads the Abstract. */
function verifyGroups(messages: string[]): { title: string; items: string[] }[] {
  const titleOf = (m: string): string => {
    if (m.startsWith('Abstract total')) return 'Abstract total'
    if (m.startsWith('Work type')) return 'By work type / category'
    if (m.startsWith('Ward')) return 'Ward wise'
    if (m.startsWith('Circle')) return 'Circle wise'
    if (m.startsWith('List of works')) return 'List of works'
    return 'Other'
  }
  const groups = new Map<string, string[]>()
  for (const m of messages) groups.set(titleOf(m), [...(groups.get(titleOf(m)) ?? []), m])
  const rank = ['Abstract total', 'By work type / category', 'Ward wise', 'Circle wise', 'List of works', 'Other']
  return [...groups.entries()]
    .sort((a, b) => rank.indexOf(a[0]) - rank.indexOf(b[0]))
    .map(([title, items]) => ({ title, items }))
}

type VerifyStatus = 'idle' | 'checking' | 'ok' | 'error'

// Same label/color per status as the Dashboard's MF_TILES (OverviewDashboard.tsx)
// so a status reads identically wherever it's shown in the app.
// Cell styling by the sheet's own heading, since the columns follow the workbook's order.
function colClass(header: string): string {
  const h = header.toLowerCase()
  if (/^sl\.?\s*no/.test(h) || h.includes('estimate')) return 'mf-works-num'
  if (h.includes('agency')) return ''
  if (h.includes('remark')) return 'mf-works-remarks'
  if (/name.*work/.test(h)) return 'mf-works-name'
  return 'mf-works-nowrap'
}

// Fixed column widths (px), chosen from the sheet's own heading — like the Works List table, so the browser
// never has to measure every cell of a few hundred rows to size the columns.
function colWidth(header: string): number {
  const h = header.toLowerCase()
  if (/^sl\.?\s*no/.test(h)) return 56
  if (/name.*work/.test(h)) return 380
  if (h.includes('remark')) return 300
  if (h.includes('agency')) return 190
  if (h.includes('estimate')) return 96
  if (h.includes('ward') || h.includes('circle')) return 150
  if (h.includes('date')) return 112
  if (h.includes('status')) return 124
  return 120
}

const STATUS_META: Record<MfStatusKey, { label: string; color: string }> = {
  completed: { label: 'Completed', color: 'var(--ok)' },
  progressTotal: { label: 'In Progress', color: 'var(--teal-500)' },
  toBeStarted: { label: 'To Be Started', color: 'var(--violet-500)' },
  tenderProcess: { label: 'Tender Process', color: 'var(--sky-500)' },
  heldUp: { label: 'Held Up', color: 'var(--warn)' },
  cancelled: { label: 'Cancelled', color: 'var(--danger)' }
}

/**
 * Lives on the Works List page, next to the works-database link import — the
 * Monitoring Format workbook is fetched the same way (paste a link or pick a
 * file) and shows this office's own progress-of-works table.
 */
export default function MonitoringFormatCard({
  office,
  monitoringFormat,
  monitoringFormatWorks,
  onImportMonitoringFormat,
  monitoringFormatLink,
  onSaveMonitoringFormatLink,
  onImportedFromLink,
  onDeleteMonitoringFormat,
  hiddenColumns = [],
  onHiddenColumnsChange,
  onExportStatusPdf,
  onExportCombinedPdf,
  onExportAbstractPdf,
  onExportFilteredWorks,
  worksLoading
}: Props) {
  const [busy, setBusy] = useState(false)
  const worksTable = useMemo(
    () => monitoringFormatWorksTable(monitoringFormatWorks ?? [], monitoringFormatWorks ?? [], hiddenColumns),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [monitoringFormatWorks, hiddenColumns.join('\u0001')]
  )
  const [pdfBusy, setPdfBusy] = useState<string | null>(null)

  // Filters on the works table itself: one per column (a list when the column has few distinct
  // values, otherwise a "contains" box), opened from the funnel on the column heading.
  const [colFilters, setColFilters] = useState<Record<number, string>>({})
  // Which per-column buttons the headings show: the filter funnels, the delete ×, or neither (set by the toolbar buttons).
  const [colMode, setColMode] = useState<'filter' | 'delete' | null>(null)
  // Filter / Delete-columns mode switches itself off on a click anywhere else in the workspace; clicks on the toggles,
  // the × / funnel buttons and the filter pop-up (all data-mf-mode-ui) are left alone, so several columns can be handled
  // in a row. Filters already applied stay applied.
  useEffect(() => {
    if (!colMode) return
    const off = (e: MouseEvent) => {
      if (!(e.target as Element | null)?.closest?.('[data-mf-mode-ui]')) {
        setColMode(null)
        setFilterCol(null)
      }
    }
    document.addEventListener('mousedown', off)
    return () => document.removeEventListener('mousedown', off)
  }, [colMode])
  const works = monitoringFormatWorks ?? []
  const cellText = (i: number, c: number): string => {
    const w = works[i]
    if (c === worksTable.statusCol && w?.statusKey) return STATUS_META[w.statusKey].label
    return worksTable.rows[i]?.[c] ?? ''
  }
  const colOptions = useMemo(() => {
    return worksTable.headers.map((h, c) => {
      if (/^sl\.?\s*no/i.test(h.trim()) || /estimate|amount|length|name.*work|remark/i.test(h)) return null // free-text / numeric: contains box or none
      const values = new Set<string>()
      for (let i = 0; i < works.length; i++) {
        const v = cellText(i, c).trim()
        if (v) values.add(v)
        if (values.size > 40) return null
      }
      return [...values].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [worksTable, works])
  const visibleIdx = useMemo(() => {
    const active = Object.entries(colFilters).filter(([, v]) => v.trim() !== '')
    const out: number[] = []
    for (let i = 0; i < works.length; i++) {
      let ok = true
      for (const [c, v] of active) {
        const text = cellText(i, Number(c))
        ok = colOptions[Number(c)] ? text.trim() === v : text.toLowerCase().includes(v.trim().toLowerCase())
        if (!ok) break
      }
      if (ok) out.push(i)
    }
    return out
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [colFilters, worksTable, works, colOptions])
  const filtersOn = Object.values(colFilters).some((v) => v.trim() !== '')
  // The column whose filter pop-up is open, with its on-screen position (fixed, so the table's scroll box can't clip it).
  const [filterCol, setFilterCol] = useState<{ col: number; x: number; y: number } | null>(null)
  const filterable = (c: number) => {
    const h = worksTable.headers[c] ?? ''
    return !!colOptions[c] || !(/^sl\.?\s*no/i.test(h.trim()) || /estimate|amount|length/i.test(h))
  }
  useEffect(() => {
    if (!filterCol) return
    const close = () => setFilterCol(null)
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && close()
    window.addEventListener('mousedown', close)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('keydown', onKey)
    }
  }, [filterCol])
  const [filteredBusy, setFilteredBusy] = useState(false)
  async function downloadFiltered() {
    if (!onExportFilteredWorks || filteredBusy || visibleIdx.length === 0) return
    setFilteredBusy(true)
    setError(null)
    try {
      const saved = await onExportFilteredWorks(
        visibleIdx.map((i) => works[i]),
        'pdf'
      )
      if (!saved) setError("Nothing was saved — the download was cancelled or the works list couldn't be read.")
    } catch (e) {
      setError(`Couldn't create the download: ${e instanceof Error ? e.message : String(e)}`)
    } finally {
      setFilteredBusy(false)
    }
  }
  async function downloadPdf(key: MfStatusKey | 'all' | 'abstract', label: string) {
    if (!onExportStatusPdf || pdfBusy) return
    setPdfBusy(label)
    setError(null)
    try {
      const saved = key === 'abstract' ? await onExportAbstractPdf?.() : await onExportStatusPdf(key, label)
      if (!saved)
        setError(
          `Couldn't create the ${label} PDF — either the save was cancelled or the works list couldn't be read. Re-import the Monitoring Format link and try again.`
        )
    } catch (e) {
      setError(`Couldn't create the ${label} PDF: ${e instanceof Error ? e.message : String(e)}`)
    } finally {
      setPdfBusy(null)
    }
  }
  const [error, setError] = useState<string | null>(null)
  const [combineOpen, setCombineOpen] = useState(false)
  const [combineKeys, setCombineKeys] = useState<MfStatusKey[]>([])
  async function downloadCombined() {
    if (!onExportCombinedPdf || pdfBusy || combineKeys.length === 0) return
    // Keep the tile order in the label regardless of the order they were ticked.
    const ordered = (Object.keys(STATUS_META) as MfStatusKey[]).filter((k) => combineKeys.includes(k))
    const names = ordered.map((k) => STATUS_META[k].label)
    // "A", "A and B", "A, B and C"
    const label = names.length > 1 ? `${names.slice(0, -1).join(', ')} and ${names[names.length - 1]}` : names[0]
    setPdfBusy(label)
    setError(null)
    try {
      const saved = await onExportCombinedPdf(ordered, label)
      if (!saved)
        setError(
          `Couldn't create the ${label} PDF — either the save was cancelled or there are no works in those statuses.`
        )
      else {
        // Downloaded: clear the ticks and fold the panel away, ready for the next combination.
        setCombineKeys([])
        setCombineOpen(false)
      }
    } catch (e) {
      setError(`Couldn't create the ${label} PDF: ${e instanceof Error ? e.message : String(e)}`)
    } finally {
      setPdfBusy(null)
    }
  }
  const [link, setLink] = useState(monitoringFormatLink ?? '')
  const [verifyStatus, setVerifyStatus] = useState<VerifyStatus>('idle')
  const [verifyIssues, setVerifyIssues] = useState<string[]>([])
  const [linkNote, setLinkNote] = useState('')
  const [verifyOpen, setVerifyOpen] = useState(false)

  // Verify results belong to the office / data they were run on: switching office, or re-importing / clearing the
  // data, drops them (the card stays mounted across office switches, so they would otherwise follow you).
  useEffect(() => {
    setVerifyStatus('idle')
    setVerifyIssues([])
    setVerifyOpen(false)
    setColMode(null)
  }, [office.corporation, office.zone, office.circle, monitoringFormat, monitoringFormatWorks])

  async function verify() {
    if (!monitoringFormat || !monitoringFormatWorks) return
    setVerifyStatus('checking')
    // Yield one tick so the "Verifying…" state actually paints before the
    // (synchronous) scan runs, instead of jumping straight to the result.
    await new Promise((resolve) => setTimeout(resolve, 0))
    const issues = findMonitoringFormatErrors(monitoringFormat, monitoringFormatWorks)
    setVerifyIssues(issues.map((v) => v.message))
    setVerifyStatus(issues.length === 0 ? 'ok' : 'error')
    setVerifyOpen(issues.length > 0)
  }

  // Re-sync the input when switching to an office with its own remembered
  // link (or none) — but not on every render, so typing isn't clobbered.
  useEffect(() => {
    setLink(monitoringFormatLink ?? '')
  }, [monitoringFormatLink, office.circle, office.zone, office.corporation])

  async function importFile() {
    setError(null)
    setBusy(true)
    try {
      const sheets = await api.pickDataSheet()
      if (!sheets) return
      onImportMonitoringFormat(
        extractMonitoringFormatForOffice(sheets, withZoneCircles(office)),
        extractMonitoringFormatWorksForOffice(sheets, withZoneCircles(office))
      )
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  async function importFromLink() {
    const url = link.trim()
    if (!url || busy) return
    setError(null)
    setLinkNote('')
    setBusy(true)
    try {
      const sheets = await api.importAllSheetsFromLink(url)
      onImportMonitoringFormat(
        extractMonitoringFormatForOffice(sheets, withZoneCircles(office)),
        extractMonitoringFormatWorksForOffice(sheets, withZoneCircles(office))
      )
      onSaveMonitoringFormatLink(url)
      setLinkNote(onImportedFromLink?.(sheets, url) || '')
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="card mf-card">
      {verifyOpen &&
        verifyStatus === 'error' &&
        createPortal(
          <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => setVerifyOpen(false))}>
            <div
              className="confirm-modal mf-verify-modal"
              role="dialog"
              aria-modal="true"
              onClick={(e) => e.stopPropagation()}
            >
              <button
                type="button"
                className="mf-verify-close"
                aria-label="Close"
                title="Close"
                onClick={() => setVerifyOpen(false)}
              >
                <IconClose />
              </button>
              <div className="confirm-ic">
                <IconWarn />
              </div>
              <h3>
                {verifyIssues.length} error{verifyIssues.length === 1 ? '' : 's'} found in the Monitoring Format
              </h3>
              <p className="confirm-hint">
                The Abstract is checked against the list of works: total works and every status, for the sheet as a
                whole, each work type and each ward / circle.
              </p>
              <div className="mf-verify-groups">
                {verifyGroups(verifyIssues).map((g) => (
                  <div key={g.title}>
                    <h4>
                      {g.title} ({g.items.length})
                    </h4>
                    <ul className="works-list-violation-list">
                      {g.items.map((m, i) => (
                        <li key={i}>{m}</li>
                      ))}
                    </ul>
                  </div>
                ))}
              </div>
              <div className="confirm-actions">
                <button className="primary" onClick={() => setVerifyOpen(false)}>
                  Close
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
      <div className="mf-body">
        <div className="mf-body-toolbar">
          {monitoringFormat && (
            <div className="verify-block">
              {works.length > 0 && (
                <button
                  type="button"
                  data-mf-mode-ui
                  className={`ghost${colMode === 'filter' ? ' mf-mode-on' : ''}${filtersOn ? ' mf-filter-applied' : ''}`}
                  aria-pressed={colMode === 'filter' || filtersOn}
                  onClick={() => {
                    setFilterCol(null)
                    // Clicking Filter again (while it is on, or while filters are applied) removes the filters.
                    if (colMode === 'filter' || filtersOn) {
                      setColFilters({})
                      setColMode(null)
                    } else {
                      setColMode('filter')
                    }
                  }}
                  title={
                    colMode === 'filter' || filtersOn
                      ? 'Remove the applied filters'
                      : 'Show a filter on each column heading'
                  }
                >
                  Filter
                </button>
              )}
              {works.length > 0 && onHiddenColumnsChange && (
                <button
                  type="button"
                  data-mf-mode-ui
                  className={`ghost${colMode === 'delete' ? ' mf-mode-on' : ''}`}
                  aria-pressed={colMode === 'delete'}
                  onClick={() => {
                    setFilterCol(null)
                    setColMode((m) => (m === 'delete' ? null : 'delete'))
                  }}
                  title="Show a delete × on each column heading (the column also leaves downloaded PDFs and Excel)"
                >
                  Delete columns
                </button>
              )}
              {works.length > 0 && onExportFilteredWorks && (
                <button
                  type="button"
                  className="ghost"
                  onClick={() => void downloadFiltered()}
                  disabled={!filtersOn || filteredBusy || visibleIdx.length === 0}
                  title={filtersOn ? 'Download the works left after filtering' : 'Apply a column filter first'}
                >
                  {filteredBusy ? 'Preparing…' : 'Download Filtered list'}
                </button>
              )}
              {hiddenColumns.length > 0 && onHiddenColumnsChange && (
                <button
                  type="button"
                  className="ghost"
                  onClick={() => {
                    setColFilters({})
                    onHiddenColumnsChange([])
                  }}
                  title="Bring back the deleted columns"
                >
                  Restore columns ({hiddenColumns.length})
                </button>
              )}
              <button
                type="button"
                className="verify-btn"
                onClick={() => void verify()}
                disabled={verifyStatus === 'checking'}
              >
                {verifyStatus === 'checking' ? 'Verifying…' : 'Verify'}
              </button>
              {verifyStatus === 'ok' && (
                <span className="verify-ok">
                  <IconCheck /> No errors found
                </span>
              )}
              {verifyStatus === 'error' && (
                <button type="button" className="verify-btn verify-blinking" onClick={() => setVerifyOpen(true)}>
                  <IconWarn /> {verifyIssues.length} error{verifyIssues.length === 1 ? '' : 's'} — view
                </button>
              )}
            </div>
          )}
          <button className="ghost mf-new-btn" onClick={importFile} disabled={busy}>
            <IconPlus /> {busy ? 'Importing…' : 'Import file'}
          </button>
          {monitoringFormat && onDeleteMonitoringFormat && (
            <button
              type="button"
              className="ghost danger-ghost"
              onClick={onDeleteMonitoringFormat}
              disabled={busy}
              title="Delete this office's Monitoring Format data"
            >
              Delete database
            </button>
          )}
        </div>
        <div className="link-import-row mf-link-row">
          <IconLink />
          <input
            value={link}
            placeholder="Paste the Monitoring Format Google Sheets link…"
            onChange={(e) => {
              setLink(e.target.value)
              setError(null)
            }}
            onKeyDown={(e) => e.key === 'Enter' && importFromLink()}
            disabled={busy}
          />
          <button className="primary" onClick={importFromLink} disabled={busy || !link.trim()}>
            {busy ? 'Fetching…' : monitoringFormat ? 'Refresh' : 'Fetch'}
          </button>
        </div>
        {error && <p className="mf-error">{error}</p>}
        {!monitoringFormat ? (
          <div className="empty">
            <p>Import the office's Monitoring Format workbook to see works progress here.</p>
          </div>
        ) : (
          <>
            <p className="mf-sub">
              {monitoringFormat.officeLabel}
              {monitoringFormat.asOfDate ? ` · As of ${monitoringFormat.asOfDate}` : ''}
            </p>
            <div className="mf-stats">
              {(
                [
                  ['Abstract', 'abstract', null],
                  ['Total Works', 'all', monitoringFormat.totals.totalWorks],
                  ['Completed', 'completed', monitoringFormat.totals.completed],
                  ['In Progress', 'progressTotal', monitoringFormat.totals.progressTotal],
                  ['To Be Started', 'toBeStarted', monitoringFormat.totals.toBeStarted],
                  ['Tender Process', 'tenderProcess', monitoringFormat.totals.tenderProcess],
                  ['Held Up', 'heldUp', monitoringFormat.totals.heldUp],
                  ['Cancelled', 'cancelled', monitoringFormat.totals.cancelled]
                ] as const
              ).map(([label, key, bucket]) => (
                <button
                  type="button"
                  key={label}
                  className={`mf-stat mf-stat-btn${bucket ? '' : ' mf-stat-abstract'}`}
                  disabled={(bucket ? bucket.no === 0 : false) || pdfBusy !== null}
                  title={
                    key === 'abstract'
                      ? 'Download the Monitoring Format Abstract as a PDF'
                      : `Download the ${label} works as a PDF`
                  }
                  onClick={() => downloadPdf(key, label)}
                >
                  <span className="mf-stat-label">{pdfBusy === label ? '…' : label}</span>
                  {bucket && (
                    <>
                      <span className="mf-stat-no">{pdfBusy === label ? '…' : bucket.no}</span>
                      <span className="mf-stat-amt">{indianDigitGroups(bucket.amt)}</span>
                    </>
                  )}
                </button>
              ))}
              {onExportCombinedPdf && (
                <button
                  type="button"
                  className={`mf-stat mf-stat-btn mf-stat-abstract${combineOpen ? ' on' : ''}`}
                  title="Pick any combination of statuses and download them together as one PDF"
                  onClick={() => setCombineOpen((o) => !o)}
                >
                  <span className="mf-stat-label">Custom</span>
                </button>
              )}
            </div>
            {combineOpen && onExportCombinedPdf && (
              <div className="mf-combine">
                <span className="mf-combine-title">Download together:</span>
                {(Object.keys(STATUS_META) as MfStatusKey[]).map((k) => (
                  <label key={k} className="mf-combine-opt">
                    <input
                      type="checkbox"
                      checked={combineKeys.includes(k)}
                      onChange={(e) =>
                        setCombineKeys((prev) => (e.target.checked ? [...prev, k] : prev.filter((x) => x !== k)))
                      }
                    />
                    {STATUS_META[k].label}
                    <span className="mf-combine-count">{monitoringFormat.totals[k].no}</span>
                  </label>
                ))}
                <button
                  type="button"
                  className="primary"
                  disabled={combineKeys.length === 0 || pdfBusy !== null}
                  onClick={downloadCombined}
                >
                  Download PDF
                </button>
              </div>
            )}
            {worksLoading && !(monitoringFormatWorks && monitoringFormatWorks.length > 0) && (
              <p className="mf-sub">Loading the works list…</p>
            )}
            {monitoringFormatWorks && monitoringFormatWorks.length > 0 && (
              <>
                <div className="mf-filter-bar">
                  <span className="mf-filter-count">
                    {visibleIdx.length} of {works.length} works
                  </span>
                </div>
                {filterCol && filterable(filterCol.col) && (
                  <div
                    data-mf-mode-ui
                    className="mf-filter-pop"
                    style={{ left: filterCol.x, top: filterCol.y }}
                    onMouseDown={(e) => e.stopPropagation()}
                  >
                    <div className="mf-filter-pop-title">{worksTable.headers[filterCol.col]}</div>
                    {colOptions[filterCol.col] ? (
                      <div className="mf-filter-pop-list">
                        {['', ...colOptions[filterCol.col]!].map((o) => (
                          <button
                            type="button"
                            key={o || '__all'}
                            className={(colFilters[filterCol.col] ?? '') === o ? 'on' : ''}
                            onClick={() => {
                              setColFilters((p) => ({ ...p, [filterCol.col]: o }))
                              setFilterCol(null)
                            }}
                          >
                            {o || 'All'}
                          </button>
                        ))}
                      </div>
                    ) : (
                      <input
                        autoFocus
                        value={colFilters[filterCol.col] ?? ''}
                        placeholder="Contains…"
                        onChange={(e) => setColFilters((p) => ({ ...p, [filterCol.col]: e.target.value }))}
                        onKeyDown={(e) => e.key === 'Enter' && setFilterCol(null)}
                      />
                    )}
                  </div>
                )}
                <div className="mf-works-wrap">
                  <table
                    className="mf-works-table"
                    style={{ tableLayout: 'fixed', width: worksTable.headers.reduce((sum, h) => sum + colWidth(h), 0) }}
                  >
                    <colgroup>
                      {worksTable.headers.map((h, c) => (
                        <col key={c} style={{ width: colWidth(h) }} />
                      ))}
                    </colgroup>
                    <thead>
                      <tr>
                        {worksTable.headers.map((h, c) => (
                          <th
                            key={c}
                            className={`mf-th${colClass(h) === 'mf-works-num' ? ' mf-works-num' : ''}${(colFilters[c] ?? '').trim() || filterCol?.col === c ? ' mf-th-active' : ''}`}
                          >
                            {h}
                            <span className="mf-th-tools">
                              {onHiddenColumnsChange && colMode === 'delete' && (
                                <button
                                  type="button"
                                  data-mf-mode-ui
                                  className="mf-th-delete"
                                  title={`Delete the ${h} column (also from downloaded PDFs and Excel)`}
                                  aria-label={`Delete the ${h} column`}
                                  onMouseDown={(e) => e.stopPropagation()}
                                  onClick={() => {
                                    setColFilters({})
                                    setFilterCol(null)
                                    onHiddenColumnsChange([...hiddenColumns, h])
                                  }}
                                >
                                  ×
                                </button>
                              )}
                              {colMode === 'filter' && filterable(c) && (
                                <button
                                  type="button"
                                  data-mf-mode-ui
                                  className={`mf-th-filter${(colFilters[c] ?? '').trim() ? ' on' : ''}${filterCol?.col === c ? ' open' : ''}`}
                                  title={`Filter ${h}`}
                                  onMouseDown={(e) => e.stopPropagation()}
                                  onClick={(e) => {
                                    const r = e.currentTarget.getBoundingClientRect()
                                    setFilterCol(
                                      filterCol?.col === c
                                        ? null
                                        : { col: c, x: Math.min(r.left, window.innerWidth - 260), y: r.bottom + 4 }
                                    )
                                  }}
                                >
                                  <svg
                                    width="12"
                                    height="12"
                                    viewBox="0 0 24 24"
                                    fill="currentColor"
                                    aria-hidden="true"
                                  >
                                    <path d="M3 4h18l-7 8v6l-4 2v-8L3 4z" />
                                  </svg>
                                </button>
                              )}
                            </span>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {visibleIdx.map((i) => {
                        const w = works[i]
                        const meta = w.statusKey ? STATUS_META[w.statusKey] : null
                        return (
                          <tr key={`${w.winCode || w.slNo}-${i}`}>
                            {worksTable.rows[i].map((cell, c) => (
                              <td key={c} className={colClass(worksTable.headers[c])}>
                                {c === worksTable.statusCol && meta ? (
                                  <span className="mf-status-pill" style={{ ['--pill-color' as string]: meta.color }}>
                                    <i className="mf-status-dot" />
                                    {meta.label}
                                  </span>
                                ) : (
                                  cell
                                )}
                              </td>
                            ))}
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </>
        )}
      </div>
    </div>
  )
}

import { useEffect, useMemo, useState } from 'react'
import type { Office } from '../office'
import type { MfStatusKey, MonitoringFormatSummary, MonitoringFormatWorkRow } from '@core/monitoringFormat'
import { extractMonitoringFormatForOffice, extractMonitoringFormatWorksForOffice, monitoringFormatWorksTable } from '@core/monitoringFormat'
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
  /** Click a status tile: save a landscape PDF of just that status's works ('all' = every work). Resolves true once saved. */
  onExportStatusPdf?: (key: MfStatusKey | 'all', label: string) => Promise<boolean>
  /** Download one PDF of the works in ANY combination of statuses the user ticked. Resolves true once saved. */
  onExportCombinedPdf?: (keys: MfStatusKey[], label: string) => Promise<boolean>
  /** Click the Abstract tile: save the sheet's Abstract (item type + ward/circle wise) as a PDF. Resolves true once saved. */
  onExportAbstractPdf?: () => Promise<boolean>
  /** The works rows are being re-read from the saved link in the background. */
  worksLoading?: boolean
}

type VerifyStatus = 'idle' | 'checking' | 'ok' | 'error'

// Same label/color per status as the Dashboard's MF_TILES (OverviewDashboard.tsx)
// so a status reads identically wherever it's shown in the app.
// Cell styling by the sheet's own heading, since the columns follow the workbook's order.
function colClass(header: string): string {
  const h = header.toLowerCase()
  if (h.startsWith('sl.no') || h.startsWith('sl no') || h.includes('estimate')) return 'mf-works-num'
  if (h.includes('agency')) return ''
  if (h.includes('remark')) return 'mf-works-remarks'
  if (/name.*work/.test(h)) return 'mf-works-name'
  return 'mf-works-nowrap'
}

// Fixed column widths (px), chosen from the sheet's own heading — like the Works List table, so the browser
// never has to measure every cell of a few hundred rows to size the columns.
function colWidth(header: string): number {
  const h = header.toLowerCase()
  if (h.startsWith('sl.no') || h.startsWith('sl no')) return 56
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
  onExportStatusPdf,
  onExportCombinedPdf,
  onExportAbstractPdf,
  worksLoading
}: Props) {
  const [busy, setBusy] = useState(false)
  const worksTable = useMemo(() => monitoringFormatWorksTable(monitoringFormatWorks ?? []), [monitoringFormatWorks])
  const [pdfBusy, setPdfBusy] = useState<string | null>(null)
  async function downloadPdf(key: MfStatusKey | 'all' | 'abstract', label: string) {
    if (!onExportStatusPdf || pdfBusy) return
    setPdfBusy(label)
    setError(null)
    try {
      const saved = key === 'abstract' ? await onExportAbstractPdf?.() : await onExportStatusPdf(key, label)
      if (!saved) setError(`Couldn't create the ${label} PDF — either the save was cancelled or the works list couldn't be read. Re-import the Monitoring Format link and try again.`)
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
      if (!saved) setError(`Couldn't create the ${label} PDF — either the save was cancelled or there are no works in those statuses.`)
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

  async function verify() {
    if (!monitoringFormat || !monitoringFormatWorks) return
    setVerifyStatus('checking')
    // Yield one tick so the "Verifying…" state actually paints before the
    // (synchronous) scan runs, instead of jumping straight to the result.
    await new Promise((resolve) => setTimeout(resolve, 0))
    const issues = findMonitoringFormatErrors(monitoringFormat, monitoringFormatWorks)
    setVerifyIssues(issues.map((v) => v.message))
    setVerifyStatus(issues.length === 0 ? 'ok' : 'error')
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
      onImportMonitoringFormat(extractMonitoringFormatForOffice(sheets, office), extractMonitoringFormatWorksForOffice(sheets, office))
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
    setBusy(true)
    try {
      const sheets = await api.importAllSheetsFromLink(url)
      onImportMonitoringFormat(extractMonitoringFormatForOffice(sheets, office), extractMonitoringFormatWorksForOffice(sheets, office))
      onSaveMonitoringFormatLink(url)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="card mf-card">
      <div className="mf-body">
        <div className="mf-body-toolbar">
          {monitoringFormat && (
            <div className="verify-block">
              <button type="button" className="verify-btn" onClick={() => void verify()} disabled={verifyStatus === 'checking'}>
                {verifyStatus === 'checking' ? 'Verifying…' : 'Verify'}
              </button>
              {verifyStatus === 'ok' && (
                <span className="verify-ok">
                  <IconCheck /> No errors found
                </span>
              )}
              {verifyStatus === 'error' && (
                <div className="verify-banner verify-blinking" role="alert">
                  <div className="verify-banner-title">
                    <IconWarn /> Errors found
                    <button type="button" className="verify-banner-close" aria-label="Dismiss" onClick={() => setVerifyStatus('idle')}>
                      <IconClose />
                    </button>
                  </div>
                  <ul>
                    {verifyIssues.map((message, i) => (
                      <li key={i}>{message}</li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          )}
          <button className="ghost mf-new-btn" onClick={importFile} disabled={busy}>
            <IconPlus /> {busy ? 'Importing…' : 'Import file'}
          </button>
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
                    title={key === 'abstract' ? 'Download the Monitoring Format Abstract as a PDF' : `Download the ${label} works as a PDF`}
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
                        onChange={(e) => setCombineKeys((prev) => (e.target.checked ? [...prev, k] : prev.filter((x) => x !== k)))}
                      />
                      {STATUS_META[k].label}
                      <span className="mf-combine-count">{monitoringFormat.totals[k].no}</span>
                    </label>
                  ))}
                  <button type="button" className="primary" disabled={combineKeys.length === 0 || pdfBusy !== null} onClick={downloadCombined}>
                    Download PDF
                  </button>
                </div>
              )}
              {worksLoading && !(monitoringFormatWorks && monitoringFormatWorks.length > 0) && <p className="mf-sub">Loading the works list…</p>}
              {monitoringFormatWorks && monitoringFormatWorks.length > 0 && (
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
                          <th key={c} className={colClass(h) === 'mf-works-num' ? 'mf-works-num' : undefined}>
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {monitoringFormatWorks.map((w, i) => {
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
              )}
            </>
          )}
        </div>
      </div>
  )
}



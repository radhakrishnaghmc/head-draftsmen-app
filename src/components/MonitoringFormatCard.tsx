import { useEffect, useState } from 'react'
import type { Office } from '../office'
import type { MfStatusKey, MonitoringFormatSummary, MonitoringFormatWorkRow } from '@core/monitoringFormat'
import { extractMonitoringFormatForOffice, extractMonitoringFormatWorksForOffice } from '@core/monitoringFormat'
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
}

type VerifyStatus = 'idle' | 'checking' | 'ok' | 'error'

// Same label/color per status as the Dashboard's MF_TILES (OverviewDashboard.tsx)
// so a status reads identically wherever it's shown in the app.
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
  onSaveMonitoringFormatLink
}: Props) {
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
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
                    ['Total Works', monitoringFormat.totals.totalWorks],
                    ['Completed', monitoringFormat.totals.completed],
                    ['In Progress', monitoringFormat.totals.progressTotal],
                    ['To Be Started', monitoringFormat.totals.toBeStarted],
                    ['Tender Process', monitoringFormat.totals.tenderProcess],
                    ['Held Up', monitoringFormat.totals.heldUp],
                    ['Cancelled', monitoringFormat.totals.cancelled]
                  ] as const
                ).map(([label, bucket]) => (
                  <div key={label} className="mf-stat">
                    <span className="mf-stat-label">{label}</span>
                    <span className="mf-stat-no">{bucket.no}</span>
                    <span className="mf-stat-amt">{indianDigitGroups(bucket.amt)}</span>
                  </div>
                ))}
              </div>
              {monitoringFormatWorks && monitoringFormatWorks.length > 0 && (
                <div className="mf-works-wrap">
                  <table className="mf-works-table">
                    <thead>
                      <tr>
                        <th className="mf-works-num">Sl.No</th>
                        <th>Circle</th>
                        <th>Ward</th>
                        <th>Name of Work</th>
                        <th>Type of Work</th>
                        <th>Win Code</th>
                        <th className="mf-works-num">Estimate (Rs. Lakhs)</th>
                        <th>Agency Details</th>
                        <th>Target Date</th>
                        <th>Status</th>
                      </tr>
                    </thead>
                    <tbody>
                      {monitoringFormatWorks.map((w, i) => {
                        const meta = w.statusKey ? STATUS_META[w.statusKey] : null
                        return (
                          <tr key={`${w.winCode || w.slNo}-${i}`}>
                            <td className="mf-works-num">{w.slNo}</td>
                            <td className="mf-works-nowrap">{w.circle}</td>
                            <td className="mf-works-nowrap">{w.ward}</td>
                            <td className="mf-works-name">{w.workName}</td>
                            <td className="mf-works-nowrap">{w.typeOfWork}</td>
                            <td className="mf-works-nowrap">{w.winCode}</td>
                            <td className="mf-works-num">{indianDigitGroups(w.estimateAmt)}</td>
                            <td>{w.agencyDetails}</td>
                            <td className="mf-works-nowrap">{w.targetDate}</td>
                            <td className="mf-works-nowrap">
                              {meta ? (
                                <span className="mf-status-pill" style={{ ['--pill-color' as string]: meta.color }}>
                                  <i className="mf-status-dot" />
                                  {meta.label}
                                </span>
                              ) : (
                                w.status
                              )}
                            </td>
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



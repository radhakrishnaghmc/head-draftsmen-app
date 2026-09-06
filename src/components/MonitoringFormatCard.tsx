import { useEffect, useState } from 'react'
import type { Office } from '../office'
import type { MonitoringFormatSummary, MonitoringFormatWorkRow } from '@core/monitoringFormat'
import { extractMonitoringFormatForOffice, extractMonitoringFormatWorksForOffice } from '@core/monitoringFormat'
import { findMonitoringFormatErrors } from '@core/worksListAgent'
import { indianDigitGroups } from '@core/worksAmounts'
import { api } from '../ipc'
import { IconPlus, IconLink, IconChevronRight, IconWarn, IconCheck, IconClose } from './Icons'

interface Props {
  office: Office
  monitoringFormat?: MonitoringFormatSummary
  monitoringFormatWorks?: MonitoringFormatWorkRow[]
  onImportMonitoringFormat: (summary: MonitoringFormatSummary, works: MonitoringFormatWorkRow[]) => void
  monitoringFormatLink?: string
  onSaveMonitoringFormatLink: (url: string) => void
}

type VerifyStatus = 'idle' | 'checking' | 'ok' | 'error'

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
  const [open, setOpen] = useState(false)
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
      <div className="card-head">
        <button
          type="button"
          className="settings-section-toggle mf-toggle"
          onClick={() => setOpen((o) => !o)}
          aria-expanded={open}
        >
          <IconChevronRight className={`settings-section-chevron ${open ? 'open' : ''}`} />
          <h3>Monitoring Format</h3>
        </button>
        {open && (
          <>
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
          </>
        )}
      </div>
      {open && (
        <>
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
              <div className="mf-table">
                <div className="mf-row mf-head">
                  <span>Status</span>
                  <span>No.</span>
                  <span>Amount</span>
                </div>
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
                  <div key={label} className="mf-row">
                    <span>{label}</span>
                    <span>{bucket.no}</span>
                    <span>{indianDigitGroups(bucket.amt)}</span>
                  </div>
                ))}
              </div>
            </>
          )}
        </>
      )}
    </div>
  )
}

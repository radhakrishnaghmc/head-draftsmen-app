import { useState } from 'react'
import { createPortal } from 'react-dom'
import { closeOnBackdropMouseDown } from '../overlayClose'
import { api } from '../ipc'
import type { VerifyDocItem, VerifyDocResult } from '../../electron/ipc-contract'
import { IconWarn, IconCheck, IconClose } from './Icons'

type Status = 'idle' | 'checking' | 'ok' | 'error'

/**
 * Reusable "Verify" button for a document-generating section (Give Intimation,
 * Work Order/Agreement, Issue Documents, ...). Re-derives what should be in
 * each already-generated document from the app's own current source data and
 * confirms it's actually there — no AI call, nothing is corrected, it only
 * reports what it finds. Errors stay up (blinking) until dismissed via the
 * close button or the next Verify run.
 */
export function VerifyButton({
  getItems,
  getExtra
}: {
  getItems: () => Promise<VerifyDocItem[]>
  /** Checks on the source figures rather than any one document. When given, the result opens as a popup (like the Works List's) instead of an inline banner. */
  getExtra?: () => Promise<{ results: VerifyDocResult[]; skipped: string[] }>
}) {
  const [status, setStatus] = useState<Status>('idle')
  const [results, setResults] = useState<VerifyDocResult[]>([])
  const [skipped, setSkipped] = useState<string[]>([])
  const [open, setOpen] = useState(false)

  async function run() {
    setStatus('checking')
    try {
      const items = await getItems()
      const docs = await api.verifyDocuments(items)
      const extra = getExtra ? await getExtra() : { results: [], skipped: [] }
      const out = [...extra.results, ...docs]
      setResults(out)
      setSkipped(extra.skipped)
      setStatus(out.every((r) => r.ok) ? 'ok' : 'error')
      setOpen(!!getExtra)
    } catch (e) {
      setResults([{ name: 'Verify', ok: false, issues: [{ field: 'Verify', message: e instanceof Error ? e.message : String(e) }] }])
      setSkipped([])
      setStatus('error')
      setOpen(!!getExtra)
    }
  }

  return (
    <div className="verify-block">
      <button type="button" className="verify-btn" onClick={() => void run()} disabled={status === 'checking'}>
        {status === 'checking' ? 'Verifying…' : 'Verify'}
      </button>
      {open &&
        (status === 'ok' || status === 'error') &&
        createPortal(
          <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => setOpen(false))}>
            <div className="confirm-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
              <div className="confirm-ic">{status === 'error' ? <IconWarn /> : <IconCheck />}</div>
              <h3>
                {status === 'error'
                  ? `${results.reduce((n, r) => n + r.issues.length, 0)} issue${results.reduce((n, r) => n + r.issues.length, 0) === 1 ? '' : 's'} found`
                  : 'No issues found'}
              </h3>
              <p className="confirm-hint">
                {status === 'error'
                  ? 'The filled documents were checked against the L-1 sheet, the Works List database and the Monitoring Format — see each item below.'
                  : 'Every check passed — estimate, ECV, contract amount, tender %, EMD/ASD, Wincode, period of completion and the Monitoring Format.'}
              </p>
              {results
                .filter((r) => !r.ok)
                .map((r) => (
                  <div key={r.name}>
                    <h4>
                      {r.name} ({r.issues.length})
                    </h4>
                    <ul className="works-list-violation-list">
                      {r.issues.map((issue, i) => (
                        <li key={i}>{issue.message}</li>
                      ))}
                    </ul>
                  </div>
                ))}
              {skipped.length > 0 && (
                <div>
                  <h4>Not checked ({skipped.length})</h4>
                  <ul className="works-list-violation-list">
                    {skipped.map((m, i) => (
                      <li key={i}>{m}</li>
                    ))}
                  </ul>
                </div>
              )}
              <div className="confirm-actions">
                <button className="primary" onClick={() => setOpen(false)}>
                  Close
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
      {status === 'ok' && (
        <span className="verify-ok">
          <IconCheck /> No errors found
        </span>
      )}
      {status === 'error' && !getExtra && (
        <div className="verify-banner verify-blinking" role="alert">
          <div className="verify-banner-title">
            <IconWarn /> Errors found
            <button type="button" className="verify-banner-close" aria-label="Dismiss" onClick={() => setStatus('idle')}>
              <IconClose />
            </button>
          </div>
          <ul>
            {results
              .filter((r) => !r.ok)
              .flatMap((r) => r.issues.map((issue, i) => ({ ...issue, doc: r.name, key: `${r.name}-${i}` })))
              .map((issue) => (
                <li key={issue.key}>
                  <strong>{issue.doc}:</strong> {issue.message}
                </li>
              ))}
          </ul>
        </div>
      )}
    </div>
  )
}

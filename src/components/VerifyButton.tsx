import { useState } from 'react'
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
export function VerifyButton({ getItems }: { getItems: () => Promise<VerifyDocItem[]> }) {
  const [status, setStatus] = useState<Status>('idle')
  const [results, setResults] = useState<VerifyDocResult[]>([])

  async function run() {
    setStatus('checking')
    try {
      const items = await getItems()
      const out = await api.verifyDocuments(items)
      setResults(out)
      setStatus(out.every((r) => r.ok) ? 'ok' : 'error')
    } catch (e) {
      setResults([{ name: 'Verify', ok: false, issues: [{ field: 'Verify', message: e instanceof Error ? e.message : String(e) }] }])
      setStatus('error')
    }
  }

  return (
    <div className="verify-block">
      <button type="button" className="verify-btn" onClick={() => void run()} disabled={status === 'checking'}>
        {status === 'checking' ? 'Verifying…' : 'Verify'}
      </button>
      {status === 'ok' && (
        <span className="verify-ok">
          <IconCheck /> No errors found
        </span>
      )}
      {status === 'error' && (
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

import { useState } from 'react'
import { api } from '../ipc'
import { IconTrash, IconWarn } from './Icons'

/** "Delete all" for an issued-documents board: asks for the account password before anything is removed. */
export default function DeleteAllButton({ count, label, loginId, onConfirmed }: { count: number; label: string; loginId: string; onConfirmed: () => void }) {
  const [open, setOpen] = useState(false)
  const [password, setPassword] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)

  function close() {
    setOpen(false)
    setPassword('')
    setError(null)
  }

  async function confirm() {
    setBusy(true)
    setError(null)
    try {
      if (await api.verifyPassword(loginId, password)) {
        onConfirmed()
        close()
      } else setError('Incorrect password.')
    } catch {
      setError('Could not verify the password. Check your internet connection and try again.')
    } finally {
      setBusy(false)
    }
  }

  return (
    <>
      <button className="danger-ghost" onClick={() => setOpen(true)} disabled={count === 0}>
        <IconTrash /> Delete all ({count})
      </button>
      {open && (
        <div className="wo-modal-overlay" role="dialog" aria-modal="true" onMouseDown={(e) => e.target === e.currentTarget && close()}>
          <div className="wo-modal" style={{ maxWidth: 420, height: 'auto' }}>
            <div className="wo-modal-body" style={{ padding: 20, display: 'grid', gap: 12 }}>
              <h3 style={{ margin: 0 }}>Delete all {label}?</h3>
              <p className="estimate-hint" style={{ margin: 0 }}>
                This removes all {count} {label} from the issued documents. Enter your password to confirm.
              </p>
              <input
                type="password"
                className="ae-input"
                placeholder="Password"
                value={password}
                autoFocus
                autoComplete="current-password"
                onChange={(e) => setPassword(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && password && !busy && void confirm()}
              />
              {error && (
                <div className="notice error">
                  <IconWarn /> {error}
                </div>
              )}
              <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button className="ghost" onClick={close} disabled={busy}>
                  Cancel
                </button>
                <button className="primary" onClick={() => void confirm()} disabled={busy || !password}>
                  {busy ? 'Checking…' : 'Delete all'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  )
}

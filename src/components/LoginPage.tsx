import { useEffect, useState } from 'react'
import { api } from '../ipc'
import { IconWarn, IconWhatsApp } from './Icons'
import appLogo from '../assets/app-logo.png'
import { MAX_CONCURRENT_SESSIONS } from '@core/sessionSlots'

interface Props {
  onSuccess: () => void
}

// Remembered across app restarts (localStorage, not sessionStorage) so the
// Login ID field is pre-filled next time — the password is never stored.
const REMEMBERED_LOGIN_ID_KEY = 'hda-remembered-login-id'

/**
 * Login gate shown before the app itself. Credentials are checked entirely
 * in the main process (against a private Google Sheet) — this component
 * only ever sees a pass/fail result, never the sheet or its contents.
 */
export default function LoginPage({ onSuccess }: Props) {
  const [loginId, setLoginId] = useState(() => localStorage.getItem(REMEMBERED_LOGIN_ID_KEY) ?? '')
  const [password, setPassword] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const [version, setVersion] = useState<string | null>(null)
  // Only offered after a normal attempt with these exact credentials has
  // already come back maxSessions — never shown up front, and cleared
  // whenever either field changes so it can't be reused with a different
  // login ID/password than the one that actually hit the limit.
  const [maxSessionsHit, setMaxSessionsHit] = useState(false)

  useEffect(() => {
    api.getAppVersion().then(setVersion)
  }, [])

  async function attemptLogin(forceLogout: boolean) {
    setBusy(true)
    setError(null)
    try {
      const result = await api.login(loginId.trim(), password, forceLogout)
      if (result.ok) {
        localStorage.setItem(REMEMBERED_LOGIN_ID_KEY, loginId.trim())
        onSuccess()
      } else if (result.maxSessions) {
        setMaxSessionsHit(true)
        setError(`Already signed in on ${MAX_CONCURRENT_SESSIONS} devices. Log out from one of them, or force-log-out every other device below.`)
      } else {
        setMaxSessionsHit(false)
        setError('Incorrect login ID or password.')
      }
    } catch {
      // Never surface the raw error here — it's almost always a network
      // failure reaching the credentials sheet, and its message (e.g. a DNS
      // error naming the host) would leak that logins are checked against a
      // Google Sheet. Show a generic, actionable message instead.
      setError('Could not verify login. Check your internet connection and try again.')
    } finally {
      setBusy(false)
    }
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault()
    if (!loginId.trim() || !password) {
      setError('Enter your login ID and password.')
      return
    }
    await attemptLogin(false)
  }

  async function forceLogout() {
    if (
      !window.confirm(
        `This will immediately sign out every other device currently logged in as "${loginId.trim()}". Continue?`
      )
    ) {
      return
    }
    await attemptLogin(true)
  }

  return (
    <div className="login-screen">
      <form className="login-card" onSubmit={submit}>
        <img src={appLogo} alt="" className="login-logo" />
        <h1>Head Draughtsman</h1>
        <p className="hint">Sign in to continue.</p>

        <label className="tender-field">
          <span>Login ID</span>
          <input
            value={loginId}
            onChange={(e) => {
              setLoginId(e.target.value)
              setError(null)
              setMaxSessionsHit(false)
            }}
            autoFocus
            autoComplete="username"
          />
        </label>

        <label className="tender-field">
          <span>Password</span>
          <input
            type="password"
            value={password}
            onChange={(e) => {
              setPassword(e.target.value)
              setError(null)
              setMaxSessionsHit(false)
            }}
            autoComplete="current-password"
          />
        </label>

        {error && (
          <div className="notice error">
            <IconWarn />
            {error}
          </div>
        )}

        <button className="primary login-submit" type="submit" disabled={busy}>
          {busy ? 'Signing in…' : 'Log In'}
        </button>

        {maxSessionsHit && (
          <button
            type="button"
            className="login-force-logout"
            disabled={busy}
            onClick={() => void forceLogout()}
          >
            {busy ? 'Signing in…' : 'Log out other devices & sign in'}
          </button>
        )}

        <p className="login-credits">
          App developed by Radhakrishna, HD{version ? ` · v${version}` : ''}
        </p>
        <button
          type="button"
          className="wa-contact"
          title="Contact Admin on WhatsApp"
          onClick={() => api.openPath('https://wa.me/919063836115?text=Hello')}
        >
          <IconWhatsApp /> Contact Admin on WhatsApp
        </button>
      </form>
    </div>
  )
}

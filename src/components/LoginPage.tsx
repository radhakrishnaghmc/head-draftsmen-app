import { useEffect, useState } from 'react'
import { api } from '../ipc'
import { IconWarn } from './Icons'
import appLogo from '../assets/app-logo.png'

interface Props {
  onSuccess: (zone?: string, circle?: string, circleNumber?: string) => void
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

  useEffect(() => {
    api.getAppVersion().then(setVersion)
  }, [])

  async function submit(e: React.FormEvent) {
    e.preventDefault()
    if (!loginId.trim() || !password) {
      setError('Enter your login ID and password.')
      return
    }
    setBusy(true)
    setError(null)
    try {
      const result = await api.login(loginId.trim(), password)
      if (result.ok) {
        localStorage.setItem(REMEMBERED_LOGIN_ID_KEY, loginId.trim())
        onSuccess(result.zone, result.circle, result.circleNumber)
      }
      else if (result.maxSessions)
        setError('Already signed in on 2 devices. Log out from one of them and try again.')
      else setError('Incorrect login ID or password.')
    } catch {
      setError('Could not verify login — check your internet connection and try again.')
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="login-screen">
      <form className="login-card" onSubmit={submit}>
        <img src={appLogo} alt="" className="login-logo" />
        <h1>Head Draftsman</h1>
        <p className="hint">Sign in to continue.</p>

        <label className="tender-field">
          <span>Login ID</span>
          <input
            value={loginId}
            onChange={(e) => {
              setLoginId(e.target.value)
              setError(null)
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

        <p className="login-credits">
          App developed by Radhakrishna, HD{version ? ` · v${version}` : ''}
        </p>
      </form>
    </div>
  )
}

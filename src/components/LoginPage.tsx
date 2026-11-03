import { useEffect, useState } from 'react'
import { api } from '../ipc'
import { IconWarn, IconEye, IconEyeOff } from './Icons'
import appLogo from '../assets/app-logo.png'
import type { LoginResult } from '@core/auth'
import { MAX_CONCURRENT_SESSIONS } from '@core/sessionSlots'
import {
  validateSignup,
  validateGoogleSignup,
  MIN_PASSWORD_LENGTH,
  type SignupRequest,
  type SignupErrors
} from '@core/signupRequest'
import { CORPORATIONS, zonesOf, circlesOf } from '../zoneCircleDirectory'

interface Props {
  /** The sign-in result: `offices` = the offices on the login (it is confined to them); `loginId` = the id its data is stored under. */
  onSuccess: (info?: Pick<LoginResult, 'offices' | 'loginId' | 'email'>) => void
}

// Remembered across app restarts (localStorage, not sessionStorage) so the
// Login ID field is pre-filled next time — the password is never stored.
// Exported so AuthGate can read the signed-in user's login ID for display
// (e.g. in the profile menu) without threading it through extra props.
export const REMEMBERED_LOGIN_ID_KEY = 'hda-remembered-login-id'
// "Remember password": the password sealed by the OS keychain (never plain text).
const REMEMBERED_PASSWORD_KEY = 'hda-remembered-password'

/**
 * Login gate shown before the app itself. Credentials are checked entirely
 * in the main process (against a private Google Sheet) — this component
 * only ever sees a pass/fail result, never the sheet or its contents.
 */
// The only Login ID usable in `npm run dev` — keeps local development off of
// real office accounts and their live per-login state, always exercising the
// same known test account instead.
const DEV_TEST_LOGIN_ID = 'radhakrishnaghmc@gmail.com'

// Set VITE_DEV_PASSWORD in a local, gitignored .env(.local) file to skip the
// login screen entirely in `npm run dev` — never set/read outside DEV, so a
// real build never auto-signs-in with anything.
const DEV_AUTOLOGIN_PASSWORD = import.meta.env.DEV
  ? (import.meta.env.VITE_DEV_PASSWORD as string | undefined)
  : undefined

/** A password box with a show/hide eye button, so what is being typed can be checked. */
function PasswordInput(props: Omit<React.InputHTMLAttributes<HTMLInputElement>, 'type'>) {
  const [shown, setShown] = useState(false)
  return (
    <span className="password-input">
      <input {...props} type={shown ? 'text' : 'password'} />
      <button
        type="button"
        className="password-toggle"
        aria-label={shown ? 'Hide password' : 'Show password'}
        title={shown ? 'Hide password' : 'Show password'}
        onClick={() => setShown((v) => !v)}
      >
        {shown ? <IconEyeOff /> : <IconEye />}
      </button>
    </span>
  )
}

// Dev auto-login fires once per launch: after a logout the form is shown, so another login can be used.
let devAutoLoginUsed = false

const EMPTY_SIGNUP: SignupRequest = {
  name: '',
  password: '',
  confirmPassword: '',
  mobile: '',
  email: '',
  corporation: '',
  zone: '',
  circle: ''
}

/** "Sign up" — creates an account with the user's own Login ID and password. The administrator approves it before it can sign in. */
function SignupForm({
  onBack,
  google
}: {
  onBack: () => void
  /** A first-time Google user: the email and password come from Google, and the account is created on submit. */
  google?: { email: string; onSignedIn: (result: LoginResult) => void }
}) {
  const [form, setForm] = useState<SignupRequest>(EMPTY_SIGNUP)
  const [errors, setErrors] = useState<SignupErrors>({})
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [done, setDone] = useState(false)
  // After step 1: the confirmation link has been emailed and we are waiting for it to be clicked.
  const [awaitingEmail, setAwaitingEmail] = useState(false)
  const [notice, setNotice] = useState<string | null>(null)

  const zones = zonesOf(form.corporation)
  const circles = circlesOf(form.corporation, form.zone)

  const patch = (changes: Partial<SignupRequest>) => {
    setForm((f) => ({ ...f, ...changes }))
    setErrors((x) => {
      const next = { ...x }
      for (const k of Object.keys(changes) as (keyof SignupRequest)[]) delete next[k]
      return next
    })
    setError(null)
  }
  const set = (k: keyof SignupRequest) => (e: React.ChangeEvent<HTMLInputElement>) => patch({ [k]: e.target.value })

  async function sendLink() {
    setBusy(true)
    setError(null)
    setNotice(null)
    try {
      const problem = await api.submitSignup(form)
      if (problem) setError(problem)
      else {
        setAwaitingEmail(true)
        setNotice(`Confirmation link sent to ${form.email.trim().toLowerCase()}.`)
      }
    } catch {
      setError('Could not send the confirmation email. Check your internet connection and try again.')
    } finally {
      setBusy(false)
    }
  }

  async function submitGoogle() {
    const found = validateGoogleSignup(form)
    setErrors(found)
    if (Object.keys(found).length > 0) return
    setBusy(true)
    setError(null)
    try {
      const result = await api.googleSignup({
        name: form.name,
        mobile: form.mobile,
        corporation: form.corporation,
        zone: form.zone,
        circle: form.circle
      })
      if (result.ok) google?.onSignedIn(result)
      else if (result.maxSessions)
        setError(`Already signed in on ${MAX_CONCURRENT_SESSIONS} devices. Log out from one of them and try again.`)
      else setError(result.error ?? 'Could not create your account.')
    } catch {
      setError('Could not create your account. Check your internet connection and try again.')
    } finally {
      setBusy(false)
    }
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault()
    if (google) return submitGoogle()
    const found = validateSignup(form)
    setErrors(found)
    if (Object.keys(found).length > 0) return
    await sendLink()
  }

  async function confirmed() {
    setBusy(true)
    setError(null)
    setNotice(null)
    try {
      const problem = await api.completeSignup(form)
      if (problem) {
        setError(problem)
      } else {
        setDone(true)
      }
    } catch {
      setError('Could not create your account. Check your internet connection and try again.')
    } finally {
      setBusy(false)
    }
  }

  if (awaitingEmail) {
    return (
      <div className="login-screen">
        <div className="login-card">
          <img src={appLogo} alt="" className="login-logo" />
          <h1>Confirm your email</h1>
          <p className="hint">
            We emailed a link to <strong>{form.email.trim().toLowerCase()}</strong>. Open it (check Spam too), then come
            back and press the button below.
          </p>
          {notice && <p className="hint">{notice}</p>}
          {error && (
            <div className="notice error">
              <IconWarn />
              {error}
            </div>
          )}
          <button className="primary login-submit" type="button" disabled={busy} onClick={() => void confirmed()}>
            {busy ? 'Checking…' : "I've confirmed my email"}
          </button>
          <button type="button" className="login-signup-link" disabled={busy} onClick={() => void sendLink()}>
            Resend the link
          </button>
          <br />
          <button
            type="button"
            className="login-signup-link"
            disabled={busy}
            onClick={() => {
              setAwaitingEmail(false)
              setError(null)
              setNotice(null)
            }}
          >
            Change my details
          </button>
        </div>
      </div>
    )
  }

  if (done) {
    return (
      <div className="login-screen">
        <div className="login-card">
          <img src={appLogo} alt="" className="login-logo" />
          <h1>Account created</h1>
          <p className="hint">
            Your email is confirmed. Your Login ID is <strong>{form.email.trim().toLowerCase()}</strong>. You can sign
            in now with this email and the password you chose.
          </p>
          <button className="primary login-submit" type="button" onClick={onBack}>
            Back to sign in
          </button>
        </div>
      </div>
    )
  }

  const field = (k: keyof SignupRequest, label: string, props: React.InputHTMLAttributes<HTMLInputElement> = {}) => (
    <label className="tender-field">
      <span>{label}</span>
      {props.type === 'password' ? (
        <PasswordInput value={form[k]} onChange={set(k)} {...props} />
      ) : (
        <input value={form[k]} onChange={set(k)} {...props} />
      )}
      {errors[k] && <span className="field-error">{errors[k]}</span>}
    </label>
  )

  return (
    <div className="login-screen">
      <form className="login-card login-card-wide" onSubmit={submit} noValidate>
        <img src={appLogo} alt="" className="login-logo" />
        <h1>{google ? 'Finish setting up' : 'Sign up'}</h1>
        <p className="hint">
          {google ? (
            <>
              Signed in with Google as <strong>{google.email}</strong>. Choose your office to create your account.
            </>
          ) : (
            'Create your account. Your email is your Login ID — choose your own password.'
          )}
        </p>

        {field('name', 'Full name', { autoFocus: true, autoComplete: 'name' })}
        {field('mobile', 'Mobile number', { type: 'tel', autoComplete: 'tel' })}
        {!google && field('email', 'Email — this is your Login ID', { type: 'email', autoComplete: 'username' })}

        <label className="tender-field">
          <span>Corporation</span>
          <select
            value={form.corporation}
            onChange={(e) => patch({ corporation: e.target.value, zone: '', circle: '' })}
          >
            <option value="">Select corporation…</option>
            {CORPORATIONS.map((c) => (
              <option key={c.name} value={c.name}>
                {c.fullName ? `${c.name} — ${c.fullName}` : c.name}
              </option>
            ))}
          </select>
          {errors.corporation && <span className="field-error">{errors.corporation}</span>}
        </label>

        <label className="tender-field">
          <span>Zone</span>
          <select
            value={form.zone}
            disabled={!form.corporation}
            onChange={(e) => patch({ zone: e.target.value, circle: '' })}
          >
            <option value="">{form.corporation && zones.length === 0 ? 'No zones listed yet' : 'Select zone…'}</option>
            {zones.map((z) => (
              <option key={z} value={z}>
                {z}
              </option>
            ))}
          </select>
          {errors.zone && <span className="field-error">{errors.zone}</span>}
        </label>

        <label className="tender-field">
          <span>Circle</span>
          <select value={form.circle} disabled={!form.zone} onChange={(e) => patch({ circle: e.target.value })}>
            <option value="">Zone level (no circle)</option>
            {circles.map((c) => (
              <option key={c.circle} value={c.circle}>
                {c.circle} (Circle {c.cno})
              </option>
            ))}
          </select>
        </label>

        {!google && field('password', 'Password', { type: 'password', autoComplete: 'new-password' })}
        {!google && field('confirmPassword', 'Confirm password', { type: 'password', autoComplete: 'new-password' })}

        {error && (
          <div className="notice error">
            <IconWarn />
            {error}
          </div>
        )}
        <button className="primary login-submit" type="submit" disabled={busy}>
          {google ? (busy ? 'Creating account…' : 'Create account') : busy ? 'Sending link…' : 'Send confirmation link'}
        </button>
        <button type="button" className="login-signup-link" disabled={busy} onClick={onBack}>
          Back to sign in
        </button>
      </form>
    </div>
  )
}

/** "Change password" — re-checks the current password, then saves the new one. Only for accounts with a password (not Google-only ones). */
function ChangePasswordForm({ initialLoginId, onBack }: { initialLoginId: string; onBack: () => void }) {
  const [loginId, setLoginId] = useState(initialLoginId)
  const [current, setCurrent] = useState('')
  const [next, setNext] = useState('')
  const [confirm, setConfirm] = useState('')
  const [error, setError] = useState<string | null>(null)
  const [done, setDone] = useState(false)
  const [busy, setBusy] = useState(false)
  // Forgot password: step 1 emails a reset link, step 2 (`sent`) takes the password chosen on the link's page.
  const [forgot, setForgot] = useState(false)
  const [sent, setSent] = useState(false)

  async function submitForgot() {
    if (!loginId.trim()) return setError('Enter your login ID (email).')
    if (sent) {
      if (next.length < MIN_PASSWORD_LENGTH)
        return setError(`Use at least ${MIN_PASSWORD_LENGTH} characters for the new password.`)
      if (next !== confirm) return setError('The new passwords do not match.')
    }
    setBusy(true)
    try {
      const reason = sent
        ? await api.finishPasswordReset(loginId.trim(), next)
        : await api.startPasswordReset(loginId.trim())
      if (reason) setError(reason)
      else if (sent) setDone(true)
      else setSent(true)
    } catch {
      setError('Could not reset the password. Check your internet connection and try again.')
    } finally {
      setBusy(false)
    }
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault()
    if (forgot) return submitForgot()
    if (!loginId.trim() || !current) return setError('Enter your login ID and current password.')
    if (next.length < MIN_PASSWORD_LENGTH)
      return setError(`Use at least ${MIN_PASSWORD_LENGTH} characters for the new password.`)
    if (next !== confirm) return setError('The new passwords do not match.')
    if (next === current) return setError('The new password must be different from the current one.')
    setBusy(true)
    try {
      const reason = await api.changePassword(loginId.trim(), current, next)
      if (reason) setError(reason)
      else setDone(true)
    } catch {
      setError('Could not change the password. Check your internet connection and try again.')
    } finally {
      setBusy(false)
    }
  }

  const input = (
    label: string,
    value: string,
    set: (v: string) => void,
    opts: { password?: boolean; autoComplete: string; autoFocus?: boolean }
  ) => (
    <label className="tender-field">
      <span>{label}</span>
      {opts.password ? (
        <PasswordInput
          value={value}
          onChange={(e) => {
            set(e.target.value)
            setError(null)
          }}
          autoComplete={opts.autoComplete}
          autoFocus={opts.autoFocus}
        />
      ) : (
        <input
          type="text"
          value={value}
          onChange={(e) => {
            set(e.target.value)
            setError(null)
          }}
          autoComplete={opts.autoComplete}
          autoFocus={opts.autoFocus}
        />
      )}
    </label>
  )

  return (
    <div className="login-screen">
      <form className="login-card" onSubmit={submit}>
        <img src={appLogo} alt="" className="login-logo" />
        <h1>{forgot ? 'Reset password' : 'Change password'}</h1>
        {done ? (
          <>
            <div className="notice">
              {forgot ? 'Your password has been reset.' : 'Your password has been changed.'} Sign in with the new
              password.
            </div>
            <button className="primary login-submit" type="button" onClick={onBack}>
              Back to sign in
            </button>
          </>
        ) : forgot ? (
          <>
            <p className="hint">
              {sent
                ? 'We emailed you a link. Open it, choose a new password there, then enter the same password below.'
                : 'Enter your login ID and we will email you a link to reset your password.'}
            </p>
            {input('Login ID (email)', loginId, setLoginId, { autoComplete: 'username', autoFocus: !initialLoginId })}
            {sent && (
              <>
                {input('New password', next, setNext, {
                  password: true,
                  autoComplete: 'new-password',
                  autoFocus: true
                })}
                {input('Confirm new password', confirm, setConfirm, { password: true, autoComplete: 'new-password' })}
              </>
            )}
            {error && (
              <div className="notice error">
                <IconWarn />
                {error}
              </div>
            )}
            <button className="primary login-submit" type="submit" disabled={busy}>
              {busy ? 'Please wait…' : sent ? 'Reset password' : 'Email me a reset link'}
            </button>
            <p className="login-signup-row">
              {sent && (
                <>
                  <button
                    type="button"
                    className="login-signup-link"
                    disabled={busy}
                    onClick={() => {
                      setSent(false)
                      setError(null)
                    }}
                  >
                    Send the link again
                  </button>
                  {' · '}
                </>
              )}
              <button type="button" className="login-signup-link" disabled={busy} onClick={onBack}>
                Back to sign in
              </button>
            </p>
          </>
        ) : (
          <>
            <p className="hint">Enter your current password, then choose a new one.</p>
            {input('Login ID (email)', loginId, setLoginId, { autoComplete: 'username', autoFocus: !initialLoginId })}
            {input('Current password', current, setCurrent, {
              password: true,
              autoComplete: 'current-password',
              autoFocus: !!initialLoginId
            })}
            {input('New password', next, setNext, { password: true, autoComplete: 'new-password' })}
            {input('Confirm new password', confirm, setConfirm, { password: true, autoComplete: 'new-password' })}
            {error && (
              <div className="notice error">
                <IconWarn />
                {error}
              </div>
            )}
            <button className="primary login-submit" type="submit" disabled={busy}>
              {busy ? 'Saving…' : 'Change password'}
            </button>
            <p className="login-signup-row">
              <button
                type="button"
                className="login-signup-link"
                disabled={busy}
                onClick={() => {
                  setForgot(true)
                  setError(null)
                }}
              >
                Forgot your password?
              </button>
              {' · '}
              <button type="button" className="login-signup-link" disabled={busy} onClick={onBack}>
                Back to sign in
              </button>
            </p>
          </>
        )}
      </form>
    </div>
  )
}

export default function LoginPage({ onSuccess }: Props) {
  const [signingUp, setSigningUp] = useState(false)
  const [changingPassword, setChangingPassword] = useState(false)
  // A Google sign-in is waiting on the browser; the user can cancel it if they closed the Google tab.
  const [googleWaiting, setGoogleWaiting] = useState(false)
  // A Google account with no account yet: the sign-up form (office, name, mobile) is shown for it.
  const [googleSetup, setGoogleSetup] = useState<string | null>(null)
  const [loginId, setLoginId] = useState(() => localStorage.getItem(REMEMBERED_LOGIN_ID_KEY) ?? '')
  const [password, setPassword] = useState('')
  const [rememberPassword, setRememberPassword] = useState(() => !!localStorage.getItem(REMEMBERED_PASSWORD_KEY))
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const [version, setVersion] = useState<string | null>(null)
  // Only offered after a normal attempt with these exact credentials has
  // already come back maxSessions — never shown up front, and cleared
  // whenever either field changes so it can't be reused with a different
  // login ID/password than the one that actually hit the limit.
  const [maxSessionsHit, setMaxSessionsHit] = useState(false)
  // Whether the attempt that hit the device limit was a Google sign-in (so "log out other devices" repeats that, not the password login).
  const [limitViaGoogle, setLimitViaGoogle] = useState(false)
  // While a dev auto-login (see DEV_AUTOLOGIN_PASSWORD) is in flight or has
  // already succeeded, the form itself is never shown — a failure (e.g. a
  // stale password in .env) drops back to the normal form, pre-filled, so it
  // can be fixed by hand instead of looping silently.
  const [devAutoLoginPending, setDevAutoLoginPending] = useState(!!DEV_AUTOLOGIN_PASSWORD && !devAutoLoginUsed)

  useEffect(() => {
    api.getAppVersion().then(setVersion)
  }, [])

  useEffect(() => {
    const sealed = localStorage.getItem(REMEMBERED_PASSWORD_KEY)
    if (!sealed) return
    void api.unsealSecret(sealed).then((plain) => {
      if (plain) setPassword((current) => current || plain)
      else localStorage.removeItem(REMEMBERED_PASSWORD_KEY)
    })
  }, [])

  useEffect(() => {
    if (!DEV_AUTOLOGIN_PASSWORD || devAutoLoginUsed) return
    devAutoLoginUsed = true
    setPassword(DEV_AUTOLOGIN_PASSWORD)
    ;(async () => {
      setBusy(true)
      try {
        const result = await api.login(DEV_TEST_LOGIN_ID, DEV_AUTOLOGIN_PASSWORD, false)
        if (result.ok) {
          localStorage.setItem(REMEMBERED_LOGIN_ID_KEY, DEV_TEST_LOGIN_ID)
          onSuccess(result)
        } else {
          setError('Dev auto-login failed — check VITE_DEV_PASSWORD in .env.')
          setDevAutoLoginPending(false)
        }
      } catch {
        setError('Could not verify login. Check your internet connection and try again.')
        setDevAutoLoginPending(false)
      } finally {
        setBusy(false)
      }
      // eslint-disable-next-line react-hooks/exhaustive-deps
    })()
  }, [])

  async function attemptLogin(forceLogout: boolean) {
    setBusy(true)
    setError(null)
    try {
      const result = await api.login(loginId.trim(), password, forceLogout)
      if (result.ok) {
        localStorage.setItem(REMEMBERED_LOGIN_ID_KEY, loginId.trim())
        const sealed = rememberPassword ? await api.sealSecret(password) : ''
        if (sealed) localStorage.setItem(REMEMBERED_PASSWORD_KEY, sealed)
        else localStorage.removeItem(REMEMBERED_PASSWORD_KEY)
        onSuccess(result)
      } else if (result.maxSessions) {
        setMaxSessionsHit(true)
        setError(
          `Already signed in on ${MAX_CONCURRENT_SESSIONS} devices. Log out from one of them, or force-log-out every other device below.`
        )
      } else {
        setMaxSessionsHit(false)
        setError('Incorrect login ID or password.')
        localStorage.removeItem(REMEMBERED_PASSWORD_KEY)
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

  async function googleLogin(forceLogout: boolean) {
    setBusy(true)
    setGoogleWaiting(true)
    setError(null)
    try {
      const result = await api.googleLogin(forceLogout)
      if (result.ok) {
        onSuccess(result)
      } else if (result.needsOffice && result.email) {
        setGoogleSetup(result.email)
      } else if (result.maxSessions) {
        setLimitViaGoogle(true)
        setMaxSessionsHit(true)
        setError(
          `Already signed in on ${MAX_CONCURRENT_SESSIONS} devices. Log out from one of them, or force-log-out every other device below.`
        )
      } else {
        setMaxSessionsHit(false)
        // A cancel the user asked for is not an error worth showing.
        if (result.error !== 'Sign-in was cancelled.') setError(result.error ?? 'Google sign-in did not complete.')
      }
    } catch {
      setError('Could not sign in with Google. Check your internet connection and try again.')
    } finally {
      setGoogleWaiting(false)
      setBusy(false)
    }
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault()
    if (!loginId.trim() || !password) {
      setError('Enter your login ID and password.')
      return
    }
    setLimitViaGoogle(false)
    await attemptLogin(false)
  }

  async function forceLogout() {
    if (
      !window.confirm(
        `This will immediately sign out every other device currently logged in as "${limitViaGoogle ? 'this Google account' : loginId.trim()}". Continue?`
      )
    ) {
      return
    }
    if (limitViaGoogle) await googleLogin(true)
    else await attemptLogin(true)
  }

  if (googleSetup) {
    return (
      <SignupForm
        onBack={() => setGoogleSetup(null)}
        google={{ email: googleSetup, onSignedIn: (result) => onSuccess(result) }}
      />
    )
  }
  if (signingUp) return <SignupForm onBack={() => setSigningUp(false)} />
  if (changingPassword)
    return <ChangePasswordForm initialLoginId={loginId.trim()} onBack={() => setChangingPassword(false)} />

  if (devAutoLoginPending) {
    return (
      <div className="login-screen">
        <div className="login-card app-boot" />
      </div>
    )
  }

  return (
    <div className="login-screen">
      <div className="login-stack">
        <form className="login-card login-card-play" onSubmit={submit}>
          <div className="login-pane-intro">
            <span className="login-mark">
              <img src={appLogo} alt="" className="login-logo" />
            </span>
            <div className="login-intro-text">
              <h1>Head Draughtsman</h1>
              <p className="hint">Tender notices, BOQs, intimations and agreements for your office in one place.</p>
            </div>
            <svg className="login-dimline" viewBox="0 0 240 24" aria-hidden="true" focusable="false">
              <path d="M2 4v16M238 4v16M2 12h236" fill="none" stroke="currentColor" strokeWidth="1.25" />
              <path d="M2 12l7-3.5v7zM238 12l-7-3.5v7z" fill="currentColor" />
            </svg>
          </div>
          <div className="login-pane-form">
            <h2 className="login-form-title">Sign in</h2>
            <label className="tender-field">
              <span>Login ID (email)</span>
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
              <PasswordInput
                value={password}
                onChange={(e) => {
                  setPassword(e.target.value)
                  setError(null)
                  setMaxSessionsHit(false)
                }}
                autoComplete="current-password"
              />
            </label>
            <label className="login-remember">
              <input
                type="checkbox"
                checked={rememberPassword}
                onChange={(e) => setRememberPassword(e.target.checked)}
              />
              <span>Remember password</span>
            </label>
            {error && (
              <div className="notice error">
                <IconWarn />
                {error}
              </div>
            )}

            {maxSessionsHit && (
              <button type="button" className="login-force-logout" disabled={busy} onClick={() => void forceLogout()}>
                {busy ? 'Signing in…' : 'Log out other devices & sign in'}
              </button>
            )}

            <div className="login-actions">
              <div className="login-actions-links">
                <button type="button" className="login-link" disabled={busy} onClick={() => setChangingPassword(true)}>
                  Change password
                </button>
                <button type="button" className="login-link" disabled={busy} onClick={() => setSigningUp(true)}>
                  Create account
                </button>
              </div>
              <button className="primary login-submit" type="submit" disabled={busy}>
                {busy ? 'Signing in…' : 'Log in'}
              </button>
            </div>

            <div className="login-or">
              <span>or</span>
            </div>
            <button
              type="button"
              className="ghost login-google"
              disabled={busy}
              onClick={() => void googleLogin(false)}
            >
              <svg width="18" height="18" viewBox="0 0 48 48" aria-hidden="true">
                <path
                  fill="#EA4335"
                  d="M24 9.5c3.5 0 6.6 1.2 9.1 3.6l6.8-6.8C35.8 2.4 30.3 0 24 0 14.6 0 6.5 5.4 2.6 13.2l7.9 6.1C12.4 13.6 17.7 9.5 24 9.5z"
                />
                <path
                  fill="#4285F4"
                  d="M46.5 24.5c0-1.6-.1-3.1-.4-4.5H24v9h12.7c-.6 3-2.3 5.5-4.8 7.2l7.5 5.8c4.4-4.1 7.1-10.1 7.1-17.5z"
                />
                <path
                  fill="#FBBC05"
                  d="M10.5 28.7c-.5-1.5-.8-3-.8-4.7s.3-3.2.8-4.7l-7.9-6.1C.9 16.4 0 20.1 0 24s.9 7.6 2.6 10.8l7.9-6.1z"
                />
                <path
                  fill="#34A853"
                  d="M24 48c6.3 0 11.7-2.1 15.6-5.7l-7.5-5.8c-2.1 1.4-4.8 2.3-8.1 2.3-6.3 0-11.6-4.1-13.5-9.8l-7.9 6.1C6.5 42.6 14.6 48 24 48z"
                />
              </svg>
              {googleWaiting ? 'Waiting for Google…' : 'Sign in with Google'}
            </button>
            {googleWaiting && (
              <p className="login-signup-row">
                Closed the Google page?{' '}
                <button type="button" className="login-signup-link" onClick={() => void api.cancelGoogleLogin()}>
                  Cancel
                </button>
              </p>
            )}
          </div>
        </form>
        <p className="login-credits">
          Developed by Radhakrishna, HD{version ? ` · v${version}` : ''}
          {' · '}
          <button
            type="button"
            className="login-link"
            title="Contact Admin on WhatsApp"
            onClick={() => api.openPath('https://wa.me/919063836115?text=Hello')}
          >
            Contact admin
          </button>
        </p>
      </div>
    </div>
  )
}

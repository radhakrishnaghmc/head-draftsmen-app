import { useEffect, useState } from 'react'
import App from './App'
import LoginPage, { REMEMBERED_LOGIN_ID_KEY } from './components/LoginPage'
import LoginBackdrop from './components/LoginBackdrop'
import UpdateBanner from './components/UpdateBanner'
import { api } from './ipc'
import { type Office, loadOffice, saveOffice, isOfficeReady, normalizeOffice } from './office'
import { getStoredTheme } from './theme'

// The signed-in flag is cleared when the app/window closes — login is required
// once per launch, not stored indefinitely.
const SESSION_KEY = 'hda-authed'

// How long the "unveil" transition (unblur the backdrop, fade the login
// card) plays before the real, authenticated App actually mounts — must
// match the CSS transition durations on .login-backdrop in styles.css.
const REVEAL_MS = 450

export default function AuthGate() {
  // Login is required on every launch, dev included — LoginPage locks the
  // Login ID to the fixed test account in dev (import.meta.env.DEV), so
  // `npm run dev` never touches a real office's live per-login state.
  const [authed, setAuthed] = useState(() => import.meta.env.DEV || sessionStorage.getItem(SESSION_KEY) === '1')
  // True for the brief window between "login accepted" and the real App
  // actually mounting — drives the unblur/fade-out transition below.
  const [revealing, setRevealing] = useState(false)

  // The signed-in user's Login ID, for display (profile menu). LoginPage
  // writes this same key to localStorage right before calling onSuccess, so
  // reading it back here (both up front and on each fresh login) always
  // reflects whoever is actually signed in.
  const [loginId, setLoginId] = useState(() => localStorage.getItem(REMEMBERED_LOGIN_ID_KEY) ?? '')

  // The chosen office (Corporation / Zone / Circle) is selected in the sidebar,
  // not derived from the login, and remembered across launches.
  const [office, setOfficeState] = useState<Office>(() => loadOffice())

  // localStorage is the fast path, but it lives per renderer-origin — a cleared
  // cache, a reinstall, or (in dev) a shifted Vite port loses it, and it never
  // follows the user to another machine. So when nothing usable is in
  // localStorage, fall back to the office saved in the on-disk / synced state
  // (written by App, independent of the renderer origin) before showing the app,
  // so a returning user is never asked to pick their office again. Starts
  // "ready" — and so skips the fallback with no delay — whenever localStorage
  // already has an office.
  const [officeReady, setOfficeReady] = useState(() => isOfficeReady(loadOffice()))

  function setOffice(next: Office) {
    setOfficeState(next)
    saveOffice(next)
  }

  // The theme (Settings → Themes) is a body class App itself applies once it
  // mounts — but App only mounts after login, so without this the login
  // screen always showed the plain default look even when Aurora/Paper was
  // the chosen theme. Applied here too, once, so it's already right by the
  // time LoginPage first renders.
  useEffect(() => {
    const theme = getStoredTheme()
    document.body.classList.toggle('theme-aurora', theme === 'aurora')
    document.body.classList.toggle('theme-paper', theme === 'paper')
    document.body.classList.toggle('theme-cyan', theme === 'cyan')
    document.body.classList.toggle('theme-meadow', theme === 'meadow')
  }, [])

  // Gated on `authed`, not just mount: the main process only points its saved-state
  // file at THIS login (state-${loginId}.json) once login succeeds — reading before
  // that would hit the generic, never-written state.json and never find the office,
  // silently defeating this whole fallback (masked on Mac, where localStorage rarely
  // needs it; the exact failure reported on Windows, where it needs it more often).
  useEffect(() => {
    if (!authed || officeReady) return
    let cancelled = false
    ;(async () => {
      try {
        const s = await api.loadState()
        if (!cancelled && s?.office && isOfficeReady(s.office)) setOffice(normalizeOffice(s.office))
      } finally {
        if (!cancelled) setOfficeReady(true)
      }
    })()
    return () => {
      cancelled = true
    }
  }, [authed, officeReady])

  if (!authed) {
    return (
      // The real app shell, static and blurred, behind the login card —
      // see LoginBackdrop's own comment. `revealing` unblurs it and fades
      // the card out for a moment before the real, authenticated App
      // swaps in (see onSuccess below), instead of a hard instant cut.
      <div className={`login-backdrop${revealing ? ' revealing' : ''}`}>
        <LoginBackdrop />
        <LoginPage
          onSuccess={() => {
            sessionStorage.setItem(SESSION_KEY, '1')
            setLoginId(localStorage.getItem(REMEMBERED_LOGIN_ID_KEY) ?? '')
            setRevealing(true)
            setTimeout(() => setAuthed(true), REVEAL_MS)
          }}
        />
        {/* After LoginPage in DOM order, not before — .login-screen is a
            fixed full-viewport box, so an earlier sibling at the same
            z-index would have its clicks swallowed by it (see styles.css). */}
        <UpdateBanner />
      </div>
    )
  }

  // Hold the app back only while we're recovering a lost office from disk (a
  // brief, one-off IPC read; skipped entirely in the common case above) so App's
  // startup runs with the right office already in hand.
  if (!officeReady) return <div className="app-boot" />

  return (
    <>
      <UpdateBanner />
      <App
        office={office}
        onOfficeChange={setOffice}
        loginId={loginId}
        onLogout={() => {
          void api.logout()
          sessionStorage.removeItem(SESSION_KEY)
          // Office is a property of this machine's Head Draughtsman, not of the
          // login session (see office.ts) — keep it in localStorage across
          // logout so it, and everything scoped to it (Document Templates,
          // Themes, per-office Works List/To Do), is still there on the next
          // login instead of resetting to defaults.
          // `revealing` stays true from the last login (never reset) — left as
          // true here, the login screen would render with opacity:0 and
          // pointer-events:none forever (see .login-backdrop.revealing CSS).
          setRevealing(false)
          setAuthed(false)
        }}
      />
    </>
  )
}

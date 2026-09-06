import { useEffect, useState } from 'react'
import App from './App'
import LoginPage, { REMEMBERED_LOGIN_ID_KEY } from './components/LoginPage'
import { api } from './ipc'
import { type Office, loadOffice, saveOffice, isOfficeReady, normalizeOffice } from './office'

// The signed-in flag is cleared when the app/window closes — login is required
// once per launch, not stored indefinitely.
const SESSION_KEY = 'hda-authed'

export default function AuthGate() {
  // Login is required on every launch, dev included — LoginPage locks the
  // Login ID to the fixed test account in dev (import.meta.env.DEV), so
  // `npm run dev` never touches a real office's live per-login state.
  const [authed, setAuthed] = useState(() => sessionStorage.getItem(SESSION_KEY) === '1')

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
      <LoginPage
        onSuccess={() => {
          sessionStorage.setItem(SESSION_KEY, '1')
          setLoginId(localStorage.getItem(REMEMBERED_LOGIN_ID_KEY) ?? '')
          setAuthed(true)
        }}
      />
    )
  }

  // Hold the app back only while we're recovering a lost office from disk (a
  // brief, one-off IPC read; skipped entirely in the common case above) so App's
  // startup runs with the right office already in hand.
  if (!officeReady) return <div className="app-boot" />

  return (
    <App
      office={office}
      onOfficeChange={setOffice}
      loginId={loginId}
      onLogout={() => {
        void api.logout()
        sessionStorage.removeItem(SESSION_KEY)
        // Office is scoped per login, not per machine: a shared machine can see
        // different people log in, each with their own office, so the outgoing
        // user's office must not leak into the next login. Clear it locally and
        // drop back to "not ready" so the fallback effect re-fires on the next
        // login and resolves THAT user's own saved office (or asks fresh if
        // they don't have one yet) instead of reusing what's still in memory.
        saveOffice({})
        setOfficeState({})
        setOfficeReady(false)
        setAuthed(false)
      }}
    />
  )
}

import { useEffect, useState } from 'react'
import App from './App'
import LoginPage from './components/LoginPage'
import { api } from './ipc'
import { type Office, loadOffice, saveOffice, isOfficeReady, normalizeOffice } from './office'

// The signed-in flag is cleared when the app/window closes — login is required
// once per launch, not stored indefinitely.
const SESSION_KEY = 'hda-authed'

export default function AuthGate() {
  // Skip the login screen in dev (`npm run dev`) only — import.meta.env.DEV
  // is Vite's own build-mode flag, false in any built/packaged output, so
  // this can't leak into a real release. Logging out during dev still shows
  // the login screen again (setAuthed(false) below), it just isn't required
  // on launch.
  const [authed, setAuthed] = useState(() => sessionStorage.getItem(SESSION_KEY) === '1' || import.meta.env.DEV)

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

  useEffect(() => {
    if (officeReady) return
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
  }, [officeReady])

  if (!authed) {
    return (
      <LoginPage
        onSuccess={() => {
          sessionStorage.setItem(SESSION_KEY, '1')
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
      onLogout={() => {
        void api.logout()
        sessionStorage.removeItem(SESSION_KEY)
        // Keep the chosen office remembered across logout — it's a property of
        // this machine's Head Draughtsman, not of the login session, so the app
        // should never re-ask for it once picked. (It can still be changed any
        // time from the sidebar's office selector.)
        setAuthed(false)
      }}
    />
  )
}

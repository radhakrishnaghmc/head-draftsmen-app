import { sameAccountOffice, type AccountOffice } from '@core/signupRequest'
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
// Whose settings this machine's localStorage currently holds. When a different account signs in, its per-office
// settings (contacts, template choices, margins…) must not carry over — two accounts can share an office name.
const OWNER_KEY = 'hda-data-owner'
// The offices a self-registered account is confined to (kept so an offline sign-in stays confined too).
const LOCKED_OFFICE_KEY = 'hda-locked-office'
// localStorage entries that belong to the machine rather than to an account.
const KEEP_ON_ACCOUNT_CHANGE = new Set([REMEMBERED_LOGIN_ID_KEY, OWNER_KEY, 'hda-last-seen-version', 'issueDocsTheme'])

function readAccountOffices(): AccountOffice[] | undefined {
  try {
    const list = JSON.parse(localStorage.getItem(LOCKED_OFFICE_KEY) ?? 'null') as AccountOffice[] | null
    return Array.isArray(list) && list.length > 0 ? list : undefined
  } catch {
    return undefined
  }
}

const toOffice = (o: AccountOffice): Office => normalizeOffice({ corporation: o.corporation, zone: o.zone, circle: o.circle || undefined })

// How long the "unveil" transition (unblur the backdrop, fade the login
// card) plays before the real, authenticated App actually mounts — must
// match the CSS transition durations on .login-backdrop in styles.css.
const REVEAL_MS = 450

export default function AuthGate() {
  // Login is required on every launch, dev included: the dev app used to skip the screen, leaving the main
  // process with no signed-in account, so anything that needs one (Evaluation Sheet sharing, cloud sync) could
  // not work there. LoginPage still auto-signs-in the test account when VITE_DEV_PASSWORD is set.
  const [authed, setAuthed] = useState(() => sessionStorage.getItem(SESSION_KEY) === '1')
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
  const [accountOffices, setAccountOffices] = useState<AccountOffice[] | undefined>(() => readAccountOffices())

  // localStorage is the fast path, but it lives per renderer-origin — a cleared
  // cache, a reinstall, or (in dev) a shifted Vite port loses it, and it never
  // follows the user to another machine. So when nothing usable is in
  // localStorage, fall back to the office saved in the on-disk / synced state
  // (written by App, independent of the renderer origin) before showing the app,
  // so a returning user is never asked to pick their office again. Starts
  // "ready" — and so skips the fallback with no delay — whenever localStorage
  // already has an office.
  const [officeReady, setOfficeReady] = useState(() => isOfficeReady(loadOffice()))

  function updateAccountOffices(list: AccountOffice[]) {
    localStorage.setItem(LOCKED_OFFICE_KEY, JSON.stringify(list))
    setAccountOffices(list)
  }
  async function addAccountOffice(o: AccountOffice): Promise<string> {
    const r = await api.addAccountOffice(o)
    if (r.error) return r.error
    updateAccountOffices(r.offices)
    return ''
  }
  async function removeAccountOffice(o: AccountOffice): Promise<string> {
    const r = await api.removeAccountOffice(o)
    if (r.error) return r.error
    updateAccountOffices(r.offices)
    return ''
  }

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
    document.body.classList.toggle('theme-play', theme === 'play')
    document.body.classList.toggle('theme-cyan', theme === 'cyan')
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
          onSuccess={(info) => {
            sessionStorage.setItem(SESSION_KEY, '1')
            // A Google login never typed a Login ID, so the remembered one would still be an old account's name.
            if (info?.email) localStorage.setItem(REMEMBERED_LOGIN_ID_KEY, info.email)

            // A different account taking over this machine starts clean: no per-office settings, no locked office.
            const owner = info?.loginId
            const previousOwner = localStorage.getItem(OWNER_KEY)
            const sameOwner = !owner || previousOwner === owner
            if (!sameOwner && (previousOwner !== null || info?.offices)) {
              for (const key of Object.keys(localStorage)) if (!KEEP_ON_ACCOUNT_CHANGE.has(key)) localStorage.removeItem(key)
            }
            if (owner) localStorage.setItem(OWNER_KEY, owner)

            // A self-registered account is confined to the offices on its login. (Offline, the stored list applies.)
            const offices = info?.offices ?? (sameOwner ? readAccountOffices() : undefined)
            if (offices && offices.length > 0) {
              localStorage.setItem(LOCKED_OFFICE_KEY, JSON.stringify(offices))
              // Open the remembered office if it is one of theirs, else their first (sign-up) office.
              const remembered = loadOffice()
              const open = offices.find((o) => sameAccountOffice(o, { corporation: remembered.corporation ?? '', zone: remembered.zone ?? '', circle: remembered.circle ?? '' }))
              setOffice(toOffice(open ?? offices[0]))
            } else {
              localStorage.removeItem(LOCKED_OFFICE_KEY)
            }
            setAccountOffices(offices && offices.length > 0 ? offices : undefined)
            setOfficeReady(true)
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
        accountOffices={accountOffices}
        onAddAccountOffice={accountOffices ? addAccountOffice : undefined}
        onRemoveAccountOffice={accountOffices ? removeAccountOffice : undefined}
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

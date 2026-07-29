import { useState } from 'react'
import App from './App'
import LoginPage from './components/LoginPage'
import { api } from './ipc'

// The signed-in flag is cleared when the app/window closes — login is required
// once per launch, not stored indefinitely.
const SESSION_KEY = 'hda-authed'
// The office identity (Zone/Circle/Circle number) is remembered across launches
// in localStorage instead — a real login refreshes it, but keeping the last one
// means the Works List's Zone/Circle auto-fill still has an identity to stamp
// after a restart and, importantly, in dev (`npm run dev`) where the login
// screen is skipped and would otherwise never capture it. Cleared on logout.
const ZONE_KEY = 'hda-zone'
const CIRCLE_KEY = 'hda-circle'
const CIRCLE_NUMBER_KEY = 'hda-circle-number'

export default function AuthGate() {
  // Skip the login screen in dev (`npm run dev`) only — import.meta.env.DEV
  // is Vite's own build-mode flag, false in any built/packaged output, so
  // this can't leak into a real release. Logging out during dev still shows
  // the login screen again (setAuthed(false) below), it just isn't required
  // on launch.
  const [authed, setAuthed] = useState(() => sessionStorage.getItem(SESSION_KEY) === '1' || import.meta.env.DEV)
  const [zone, setZone] = useState(() => localStorage.getItem(ZONE_KEY) ?? undefined)
  const [circle, setCircle] = useState(() => localStorage.getItem(CIRCLE_KEY) ?? undefined)
  const [circleNumber, setCircleNumber] = useState(() => localStorage.getItem(CIRCLE_NUMBER_KEY) ?? undefined)

  if (!authed) {
    return (
      <LoginPage
        onSuccess={(z, c, cno) => {
          sessionStorage.setItem(SESSION_KEY, '1')
          // Remember (or clear) the office identity across launches.
          if (z) localStorage.setItem(ZONE_KEY, z)
          else localStorage.removeItem(ZONE_KEY)
          if (c) localStorage.setItem(CIRCLE_KEY, c)
          else localStorage.removeItem(CIRCLE_KEY)
          if (cno) localStorage.setItem(CIRCLE_NUMBER_KEY, cno)
          else localStorage.removeItem(CIRCLE_NUMBER_KEY)
          setZone(z)
          setCircle(c)
          setCircleNumber(cno)
          setAuthed(true)
        }}
      />
    )
  }

  return (
    <App
      loginZone={zone}
      loginCircle={circle}
      loginCircleNumber={circleNumber}
      onLogout={() => {
        void api.logout()
        sessionStorage.removeItem(SESSION_KEY)
        localStorage.removeItem(ZONE_KEY)
        localStorage.removeItem(CIRCLE_KEY)
        localStorage.removeItem(CIRCLE_NUMBER_KEY)
        setZone(undefined)
        setCircle(undefined)
        setCircleNumber(undefined)
        setAuthed(false)
      }}
    />
  )
}

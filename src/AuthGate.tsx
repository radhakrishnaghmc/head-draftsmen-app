import { useState } from 'react'
import App from './App'
import LoginPage from './components/LoginPage'
import { api } from './ipc'
import { type Office, loadOffice, saveOffice } from './office'

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

  function setOffice(next: Office) {
    setOfficeState(next)
    saveOffice(next)
  }

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

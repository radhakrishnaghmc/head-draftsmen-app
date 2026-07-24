import { useState } from 'react'
import App from './App'
import LoginPage from './components/LoginPage'
import { api } from './ipc'

// Cleared when the app/window closes — login is required once per launch,
// not stored indefinitely.
const SESSION_KEY = 'hda-authed'
const ZONE_KEY = 'hda-zone'
const CIRCLE_KEY = 'hda-circle'
const CIRCLE_NUMBER_KEY = 'hda-circle-number'

export default function AuthGate() {
  const [authed, setAuthed] = useState(() => sessionStorage.getItem(SESSION_KEY) === '1')
  const [zone, setZone] = useState(() => sessionStorage.getItem(ZONE_KEY) ?? undefined)
  const [circle, setCircle] = useState(() => sessionStorage.getItem(CIRCLE_KEY) ?? undefined)
  const [circleNumber, setCircleNumber] = useState(() => sessionStorage.getItem(CIRCLE_NUMBER_KEY) ?? undefined)

  if (!authed) {
    return (
      <LoginPage
        onSuccess={(z, c, cno) => {
          sessionStorage.setItem(SESSION_KEY, '1')
          if (z) sessionStorage.setItem(ZONE_KEY, z)
          if (c) sessionStorage.setItem(CIRCLE_KEY, c)
          if (cno) sessionStorage.setItem(CIRCLE_NUMBER_KEY, cno)
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
        sessionStorage.removeItem(ZONE_KEY)
        sessionStorage.removeItem(CIRCLE_KEY)
        sessionStorage.removeItem(CIRCLE_NUMBER_KEY)
        setAuthed(false)
      }}
    />
  )
}

import { useEffect, useState } from 'react'
import { api } from '../ipc'
import { IconRefresh } from './Icons'

/**
 * A new version downloads silently in the background (see
 * electron/autoUpdate.ts) — this just tells the user once it's ready. If
 * dismissed, the update still installs automatically the next time the app
 * quits, so nothing is lost by not restarting right away.
 */
export default function UpdateBanner() {
  const [ready, setReady] = useState(false)
  const [dismissed, setDismissed] = useState(false)

  useEffect(() => api.onUpdateDownloaded(() => setReady(true)), [])

  if (!ready || dismissed) return null

  return (
    <div className="update-banner">
      <IconRefresh />
      <span>A new version has been downloaded.</span>
      <button className="primary" onClick={() => api.restartToUpdate()}>
        Restart Now
      </button>
      <button className="ghost" onClick={() => setDismissed(true)}>
        Later
      </button>
    </div>
  )
}

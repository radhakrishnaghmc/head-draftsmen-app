import { useEffect, useState } from 'react'
import { api } from '../ipc'
import { IconRefresh, IconWarn } from './Icons'

const RELEASES_URL = 'https://github.com/radhakrishnaghmc/head-draftsmen-app/releases/latest'

/**
 * A new version downloads silently in the background (see
 * electron/autoUpdate.ts) — this just tells the user once it's ready. If
 * dismissed, the update still installs automatically the next time the app
 * quits, so nothing is lost by not restarting right away.
 *
 * Installing can fail right away on an unsigned macOS build (Squirrel.Mac
 * refuses the code signature) — without onUpdateInstallError, clicking
 * "Restart Now" would just silently do nothing, so that failure is
 * surfaced here with a manual-download fallback instead.
 */
export default function UpdateBanner() {
  const [ready, setReady] = useState(false)
  const [dismissed, setDismissed] = useState(false)
  const [installError, setInstallError] = useState<string | null>(null)

  useEffect(() => api.onUpdateDownloaded(() => setReady(true)), [])
  useEffect(() => api.onUpdateInstallError((message) => setInstallError(message)), [])

  if (installError) {
    return (
      <div className="update-banner error">
        <IconWarn />
        <span>Couldn't install the update automatically ({installError}).</span>
        <button className="primary" onClick={() => api.openPath(RELEASES_URL)}>
          Download Manually
        </button>
        <button className="ghost" onClick={() => setInstallError(null)}>
          Dismiss
        </button>
      </div>
    )
  }

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

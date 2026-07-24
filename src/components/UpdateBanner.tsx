import { useEffect, useState } from 'react'
import { api } from '../ipc'
import type { UpdateProgress } from '../../electron/ipc-contract'
import { IconRefresh, IconWarn } from './Icons'

const RELEASES_URL = 'https://github.com/radhakrishnaghmc/head-draftsmen-app/releases/latest'

/** 45230000 -> "43.1 MB" */
function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 MB'
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

/**
 * A new version downloads silently in the background (see
 * electron/autoUpdate.ts) — this shows progress while it downloads, then
 * lets the user restart once it's ready. If dismissed, the update still
 * installs automatically the next time the app quits, so nothing is lost
 * by not restarting right away.
 *
 * Installing can fail right away on an unsigned macOS build (Squirrel.Mac
 * refuses the code signature) — without onUpdateInstallError, clicking
 * "Restart Now" would just silently do nothing, so that failure is
 * surfaced here with a manual-download fallback instead.
 */
export default function UpdateBanner() {
  const [progress, setProgress] = useState<UpdateProgress | null>(null)
  const [ready, setReady] = useState(false)
  const [dismissed, setDismissed] = useState(false)
  const [installError, setInstallError] = useState<string | null>(null)

  useEffect(() => api.onUpdateProgress((p) => setProgress(p)), [])
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

  if (ready) {
    if (dismissed) return null
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

  if (progress) {
    return (
      <div className="update-banner downloading">
        <IconRefresh className="spin" />
        <div className="update-progress-body">
          <span>
            Downloading update… {formatBytes(progress.transferred)} / {formatBytes(progress.total)} (
            {Math.round(progress.percent)}%)
          </span>
          <div className="update-progress-track">
            <div className="update-progress-fill" style={{ width: `${Math.min(100, progress.percent)}%` }} />
          </div>
        </div>
      </div>
    )
  }

  return null
}

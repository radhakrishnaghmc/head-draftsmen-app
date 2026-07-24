// Silent background updates: checks on launch (and periodically while the
// app stays open), downloads a newer version automatically if found, and
// lets the renderer show a "Restart to update" prompt once it's ready — the
// user never has to manually check or re-download an installer themselves.
// Only runs in a packaged build (electron-builder's `publish` config feeds
// autoUpdater the update feed URL via the bundled app-update.yml); dev runs
// never check, since there's no packaged feed to read.
import { app, BrowserWindow } from 'electron'
import { autoUpdater } from 'electron-updater'
import { IPC } from './ipc-contract'
import type { ManualCheckResult } from './ipc-contract'

const RECHECK_INTERVAL_MS = 4 * 60 * 60 * 1000 // 4 hours — covers long-running sessions

export function initAutoUpdate(getWindow: () => BrowserWindow | null): void {
  if (!app.isPackaged) return

  autoUpdater.autoDownload = true

  const notifyDownloaded = () => {
    const win = getWindow()
    if (win && !win.isDestroyed()) {
      win.webContents.send(IPC.updateDownloaded)
    }
  }

  autoUpdater.on('update-downloaded', notifyDownloaded)
  autoUpdater.on('error', (e) => {
    // Non-fatal: the app keeps running on the current version if a check or
    // download fails (offline, host unreachable, etc.) — never block startup.
    console.error('autoUpdater error', e)
  })

  const check = () => {
    autoUpdater.checkForUpdates().catch((e) => console.error('checkForUpdates failed', e))
  }
  check()
  setInterval(check, RECHECK_INTERVAL_MS)
}

/**
 * Applies a downloaded update and restarts — call only after
 * update-downloaded has fired. If install fails right away (e.g. macOS
 * Squirrel.Mac refusing an unsigned app's code signature — a real,
 * currently-unresolved limitation on macOS specifically, not a network
 * hiccup) the app never quits and nothing else would tell the user why
 * — `onError` reports that one failure so the UI isn't just silent.
 */
export function restartToUpdate(onError: (message: string) => void): void {
  const onErr = (e: Error) => {
    autoUpdater.removeListener('error', onErr)
    onError(e.message || 'The update could not be installed automatically.')
  }
  autoUpdater.once('error', onErr)
  autoUpdater.quitAndInstall()
  // If the app hasn't quit and no error fired within a few seconds, stop
  // waiting so a later, unrelated background-check error doesn't get
  // mistakenly attributed to this restart attempt.
  setTimeout(() => autoUpdater.removeListener('error', onErr), 10_000)
}

/**
 * Triggered by the small update icon in the sidebar — same underlying check
 * as the automatic one, just with an immediate result so the UI can show
 * "Up to date" / "Update found" instead of waiting silently.
 */
export function checkForUpdatesManually(): Promise<ManualCheckResult> {
  if (!app.isPackaged) return Promise.resolve('dev-mode')

  return new Promise((resolve) => {
    let settled = false
    const finish = (result: ManualCheckResult) => {
      if (settled) return
      settled = true
      autoUpdater.removeListener('update-available', onAvailable)
      autoUpdater.removeListener('update-not-available', onNotAvailable)
      autoUpdater.removeListener('error', onError)
      resolve(result)
    }
    const onAvailable = () => finish('update-available')
    const onNotAvailable = () => finish('up-to-date')
    const onError = () => finish('error')

    autoUpdater.once('update-available', onAvailable)
    autoUpdater.once('update-not-available', onNotAvailable)
    autoUpdater.once('error', onError)
    autoUpdater.checkForUpdates().catch(() => finish('error'))
  })
}

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
import type { ManualCheckResult, UpdateProgress } from './ipc-contract'

const RECHECK_INTERVAL_MS = 4 * 60 * 60 * 1000 // 4 hours — covers long-running sessions

export function initAutoUpdate(getWindow: () => BrowserWindow | null): void {
  if (!app.isPackaged) return

  autoUpdater.autoDownload = true
  // Once a download completes, the app is sitting in "waiting to install"
  // limbo — any autoUpdater error from that point on (whether triggered by
  // the user clicking Restart Now, or by Squirrel.Mac automatically
  // re-attempting/re-validating the cached update on the next launch,
  // which happens with no user action at all) is almost certainly that
  // install itself failing, not a routine background-check hiccup, so it's
  // worth surfacing. Before a download has completed, a check/download
  // error (offline, host unreachable) stays silent — that's routine and
  // shouldn't alarm the user every time the network blips.
  let downloaded = false

  const notifyDownloaded = () => {
    downloaded = true
    const win = getWindow()
    if (win && !win.isDestroyed()) {
      win.webContents.send(IPC.updateDownloaded)
    }
  }

  const notifyProgress = (progress: UpdateProgress) => {
    const win = getWindow()
    if (win && !win.isDestroyed()) {
      win.webContents.send(IPC.updateProgress, {
        percent: progress.percent,
        transferred: progress.transferred,
        total: progress.total,
        bytesPerSecond: progress.bytesPerSecond
      })
    }
  }

  autoUpdater.on('update-downloaded', notifyDownloaded)
  autoUpdater.on('download-progress', notifyProgress)
  autoUpdater.on('error', (e) => {
    console.error('autoUpdater error', e)
    if (downloaded) {
      const win = getWindow()
      if (win && !win.isDestroyed()) {
        win.webContents.send(IPC.updateInstallError, e.message || 'The update could not be installed automatically.')
      }
    }
  })

  const check = () => {
    autoUpdater.checkForUpdates().catch((e) => console.error('checkForUpdates failed', e))
  }
  check()
  setInterval(check, RECHECK_INTERVAL_MS)
}

/** Applies a downloaded update and restarts — call only after update-downloaded has fired. */
export function restartToUpdate(): void {
  autoUpdater.quitAndInstall()
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

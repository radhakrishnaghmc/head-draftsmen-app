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

/** Applies a downloaded update and restarts — call only after update-downloaded has fired. */
export function restartToUpdate(): void {
  autoUpdater.quitAndInstall()
}

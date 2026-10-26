// Local activity log: an append-only record of what this login did in the app
// (saves/edits, sign-ins, backups, restores), kept per login next to the
// workspace cache in userData. It travels inside backups (core/backup.ts) so
// the history survives a move to another device.
import * as fs from 'fs'
import * as path from 'path'
import * as crypto from 'crypto'
import { app } from 'electron'
import { mergeActivityLogs, type ActivityEntry } from '../core/backup'

const MAX_ENTRIES = 5000

const logFile = (loginId: string) => path.join(app.getPath('userData'), `activity-${loginId}.json`)

export function readActivityLog(loginId: string): ActivityEntry[] {
  try {
    const parsed = JSON.parse(fs.readFileSync(logFile(loginId), 'utf8'))
    return Array.isArray(parsed) ? (parsed as ActivityEntry[]) : []
  } catch {
    return []
  }
}

function writeActivityLog(loginId: string, entries: ActivityEntry[]): void {
  try {
    fs.writeFileSync(logFile(loginId), JSON.stringify(entries.slice(-MAX_ENTRIES)), 'utf8')
  } catch {
    // Best-effort: a failed log write must never break the action being logged.
  }
}

export function logActivity(loginId: string | null, action: string, detail: string): void {
  if (!loginId) return
  const entry: ActivityEntry = { id: `${Date.now()}-${crypto.randomBytes(3).toString('hex')}`, ts: Date.now(), action, detail }
  writeActivityLog(loginId, [...readActivityLog(loginId), entry])
}

/** Folds a backup's history into this device's log (nothing is lost on either side). */
export function mergeIntoActivityLog(loginId: string, incoming: ActivityEntry[]): void {
  writeActivityLog(loginId, mergeActivityLogs(readActivityLog(loginId), incoming))
}

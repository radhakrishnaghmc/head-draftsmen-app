// Lets the SAME person who has already proven their password against the
// real credentials sheet at least once on THIS device log in again later
// with no internet connection — never a blanket offline bypass, and never a
// different login ID than the one that was actually verified online here.
// The password itself is never stored, only a salted scrypt hash of it (see
// core/offlineAuthCore.ts) — a leaked/stolen copy of this file can't be
// trivially reversed back into a usable password.
import * as fs from 'fs'
import * as path from 'path'
import { app } from 'electron'
import { normalizeId } from './firebaseSync'
import { makeOfflineAuthRecord, matchesOfflineAuthRecord, type OfflineAuthRecord } from '../core/offlineAuthCore'

type OfflineAuthStore = Record<string, OfflineAuthRecord>

function offlineAuthFile(): string {
  return path.join(app.getPath('userData'), 'offline-auth.json')
}

function readStore(): OfflineAuthStore {
  try {
    return JSON.parse(fs.readFileSync(offlineAuthFile(), 'utf8'))
  } catch {
    return {}
  }
}

/** Call after every SUCCESSFUL online login — refreshes this user's offline-login record on this device. */
export function rememberOfflineAuth(loginId: string, password: string): void {
  const store = readStore()
  store[normalizeId(loginId)] = makeOfflineAuthRecord(password)
  try {
    fs.writeFileSync(offlineAuthFile(), JSON.stringify(store), 'utf8')
  } catch {
    // Best-effort — a failure here just means offline login won't be
    // available next time; it must never block the (already-succeeded)
    // online login itself.
  }
}

/** Whether this exact loginId/password pair matches what was last verified online, on this device, for this specific user. */
export function checkOfflineAuth(loginId: string, password: string): boolean {
  const record = readStore()[normalizeId(loginId)]
  return !!record && matchesOfflineAuthRecord(password, record)
}

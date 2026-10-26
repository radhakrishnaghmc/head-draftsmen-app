// Encrypted whole-workspace backup (Settings > Backup / Restore).
//
// The file is AES-256-GCM sealed with a key derived (scrypt) from the login ID
// AND password, so a backup only opens under the login that made it — on any
// device, but not under another login, and not with a changed password. The
// login ID is also stored inside the encrypted payload and re-checked on
// restore, so a file can never be applied to a different account even if the
// key somehow matched. Pure Node crypto: no Electron, unit-testable.
import * as crypto from 'crypto'
import type { PersistedState } from './types'

export const BACKUP_EXTENSION = 'dhbak'
const MAGIC = 'DRAFTSMEN-BACKUP'
const FORMAT_VERSION = 1

export interface ActivityEntry {
  id: string
  /** Epoch ms. */
  ts: number
  /** Short machine-ish category, e.g. "save", "login", "backup". */
  action: string
  /** Human-readable one-liner. */
  detail: string
}

export interface BackupPayload {
  appVersion: string
  createdAt: number
  loginId: string
  state: PersistedState
  activityLog: ActivityEntry[]
}

interface BackupEnvelope {
  magic: string
  v: number
  salt: string
  iv: string
  tag: string
  data: string
}

export class BackupError extends Error {
  constructor(
    public code: 'NOT_A_BACKUP' | 'WRONG_LOGIN' | 'CORRUPT',
    message: string
  ) {
    super(message)
  }
}

const normalizeLogin = (loginId: string) => loginId.trim().toLowerCase()

function deriveKey(loginId: string, password: string, salt: Buffer): Buffer {
  return crypto.scryptSync(`${normalizeLogin(loginId)}\u0000${password}`, salt, 32)
}

export function encryptBackup(payload: BackupPayload, loginId: string, password: string): Buffer {
  const salt = crypto.randomBytes(16)
  const iv = crypto.randomBytes(12)
  const cipher = crypto.createCipheriv('aes-256-gcm', deriveKey(loginId, password, salt), iv)
  const data = Buffer.concat([cipher.update(JSON.stringify(payload), 'utf8'), cipher.final()])
  const envelope: BackupEnvelope = {
    magic: MAGIC,
    v: FORMAT_VERSION,
    salt: salt.toString('base64'),
    iv: iv.toString('base64'),
    tag: cipher.getAuthTag().toString('base64'),
    data: data.toString('base64')
  }
  return Buffer.from(JSON.stringify(envelope), 'utf8')
}

export function decryptBackup(file: Buffer, loginId: string, password: string): BackupPayload {
  let envelope: BackupEnvelope
  try {
    envelope = JSON.parse(file.toString('utf8')) as BackupEnvelope
  } catch {
    throw new BackupError('NOT_A_BACKUP', 'This is not a Head Draftsmen backup file.')
  }
  if (envelope?.magic !== MAGIC || envelope.v !== FORMAT_VERSION) {
    throw new BackupError('NOT_A_BACKUP', 'This is not a Head Draftsmen backup file.')
  }
  let plain: Buffer
  try {
    const decipher = crypto.createDecipheriv(
      'aes-256-gcm',
      deriveKey(loginId, password, Buffer.from(envelope.salt, 'base64')),
      Buffer.from(envelope.iv, 'base64')
    )
    decipher.setAuthTag(Buffer.from(envelope.tag, 'base64'))
    plain = Buffer.concat([decipher.update(Buffer.from(envelope.data, 'base64')), decipher.final()])
  } catch {
    throw new BackupError(
      'WRONG_LOGIN',
      'This backup belongs to a different login (or the password was changed since it was made), so it cannot be restored here.'
    )
  }
  let payload: BackupPayload
  try {
    payload = JSON.parse(plain.toString('utf8')) as BackupPayload
  } catch {
    throw new BackupError('CORRUPT', 'The backup file is damaged.')
  }
  if (normalizeLogin(payload.loginId ?? '') !== normalizeLogin(loginId) || !payload.state) {
    throw new BackupError('WRONG_LOGIN', 'This backup belongs to a different login and cannot be restored here.')
  }
  return payload
}

/** Union of two activity logs (restore keeps this device's history too), de-duplicated by id and oldest first. */
export function mergeActivityLogs(a: ActivityEntry[], b: ActivityEntry[]): ActivityEntry[] {
  const byId = new Map<string, ActivityEntry>()
  for (const e of [...a, ...b]) byId.set(e.id, e)
  return [...byId.values()].sort((x, y) => x.ts - y.ts)
}

const ARRAY_OR_MAP_LABEL: Record<string, string> = {
  tables: 'Works List',
  tablesByOffice: 'Works List (offices)',
  todos: 'To Do',
  mbScrutiny: 'MB Scrutiny',
  createdDocuments: 'Issue Documents',
  tenderConditions: 'Tender conditions',
  tenderReminders: 'Tender reminders',
  pendingNotices: 'Pending notices',
  bidDocumentBatches: 'Bid Document batches',
  qcParties: 'QC agencies',
  worksListLinks: 'Works List links',
  monitoringFormatByOffice: 'Monitoring Format',
  monitoringFormatWorksByOffice: 'Monitoring Format works',
  monitoringFormatLinks: 'Monitoring Format links',
  resolution: 'Collision resolutions',
  office: 'Office',
  contacts: 'Contact details',
  lastGoogleLink: 'Works List link'
}

const sizeOf = (v: unknown): number | null =>
  Array.isArray(v) ? v.length : v && typeof v === 'object' ? Object.keys(v).length : null

/**
 * Human-readable summary of what changed between two saves, from the per-key
 * JSON signatures — e.g. "To Do 5→6, Issue Documents edited". Empty when
 * nothing user-visible changed.
 */
export function describeStateChange(prevSigs: Record<string, string>, next: PersistedState): { sigs: Record<string, string>; detail: string } {
  const sigs: Record<string, string> = {}
  const parts: string[] = []
  for (const [key, label] of Object.entries(ARRAY_OR_MAP_LABEL)) {
    const value = (next as unknown as Record<string, unknown>)[key]
    const sig = JSON.stringify(value ?? null)
    sigs[key] = sig
    if (!(key in prevSigs) || prevSigs[key] === sig) continue
    const before = sizeOf(JSON.parse(prevSigs[key]))
    const after = sizeOf(value)
    parts.push(before !== null && after !== null && before !== after ? `${label} ${before}→${after}` : `${label} edited`)
  }
  return { sigs, detail: parts.join(', ') }
}

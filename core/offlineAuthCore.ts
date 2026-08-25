// Pure salted-hash logic behind offline login (see electron/offlineAuth.ts,
// which adds the Electron-only userData file I/O around this). Split out so
// the actual hashing/verification can be unit tested without an Electron
// runtime — this file only needs Node's crypto module, same as
// core/googleImport.ts already assumes plain Node (http/https) is available.
import * as crypto from 'crypto'

export interface OfflineAuthRecord {
  salt: string
  hash: string
}

function hashPassword(password: string, salt: Buffer): Buffer {
  return crypto.scryptSync(password, salt, 64)
}

/** A fresh salted record for a just-verified password — store this, never the password itself. */
export function makeOfflineAuthRecord(password: string): OfflineAuthRecord {
  const salt = crypto.randomBytes(16)
  return { salt: salt.toString('hex'), hash: hashPassword(password, salt).toString('hex') }
}

/** Whether `password` is the one `record` was made from. Constant-time compare against timing attacks. */
export function matchesOfflineAuthRecord(password: string, record: OfflineAuthRecord): boolean {
  const salt = Buffer.from(record.salt, 'hex')
  const expected = Buffer.from(record.hash, 'hex')
  const actual = hashPassword(password, salt)
  return actual.length === expected.length && crypto.timingSafeEqual(actual, expected)
}

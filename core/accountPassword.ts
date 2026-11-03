import { randomBytes, scryptSync, timingSafeEqual } from 'crypto'

/** Salted scrypt hash of a password, as "salt:hash" in hex — what is stored for self-registered accounts instead of the password. Main process only. */
export function hashPassword(password: string): string {
  const salt = randomBytes(16)
  return `${salt.toString('hex')}:${scryptSync(password, salt, 32).toString('hex')}`
}

export function verifyPassword(password: string, stored: string): boolean {
  const [saltHex, hashHex] = stored.split(':')
  if (!saltHex || !hashHex) return false
  const expected = Buffer.from(hashHex, 'hex')
  const actual = scryptSync(password, Buffer.from(saltHex, 'hex'), expected.length)
  return actual.length === expected.length && timingSafeEqual(actual, expected)
}

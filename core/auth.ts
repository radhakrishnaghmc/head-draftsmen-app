import { importTableFromGoogleLink } from './googleImport'

// Not exported, and never sent to the renderer — the sheet holding login
// credentials. Kept private to this module so the link is never visible
// outside the main process.
const LOGIN_SHEET_URL =
  'https://docs.google.com/spreadsheets/d/1-cWNTql8wN5Ii7Vx6NZ7Gl74z9yeMzZcZxhjd6JYPm8/edit?usp=sharing'

export interface LoginResult {
  ok: boolean
  /** The matched row's Zone/Circle, when present — the logged-in Head Draftsman's own office identity, shown in the sidebar. */
  zone?: string
  circle?: string
  /** The matched row's Circle number, when present — used to auto-fill the Works List's CNO column, same as Zone/Circle. */
  circleNumber?: string
  /** True when the credentials were correct but this account already has 2 other devices signed in. */
  maxSessions?: boolean
}

/**
 * Check a login attempt against the credentials sheet. Fetches fresh every
 * call (no caching) so newly added/changed credentials take effect
 * immediately. Only Zone/Circle (never the password, any other column, or
 * the sheet's contents/URL) cross back to the renderer alongside the
 * pass/fail result.
 */
export async function validateLogin(loginId: string, password: string): Promise<LoginResult> {
  const table = await importTableFromGoogleLink(LOGIN_SHEET_URL)
  const idHeader = table.headers.find((h) => /login\s*id|user/i.test(h))
  const pwHeader = table.headers.find((h) => /password/i.test(h))
  const zoneHeader = table.headers.find((h) => /^zone$/i.test(h.trim()))
  const circleHeader = table.headers.find((h) => /^circle$/i.test(h.trim()))
  const circleNumberHeader = table.headers.find((h) => /^circle\s*number$/i.test(h.trim()))
  if (!idHeader || !pwHeader) {
    throw new Error('Could not read the login credentials sheet.')
  }

  const id = loginId.trim().toLowerCase()
  const row = table.rows.find(
    (r) => (r[idHeader] ?? '').trim().toLowerCase() === id && (r[pwHeader] ?? '') === password
  )
  if (!row) return { ok: false }

  return {
    ok: true,
    zone: zoneHeader ? row[zoneHeader]?.trim() || undefined : undefined,
    circle: circleHeader ? row[circleHeader]?.trim() || undefined : undefined,
    circleNumber: circleNumberHeader ? row[circleNumberHeader]?.trim() || undefined : undefined
  }
}

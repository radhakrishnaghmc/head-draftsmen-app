import { importTableFromGoogleLink } from './googleImport'

// Not exported, and never sent to the renderer — the sheet holding login
// credentials. Kept private to this module so the link is never visible
// outside the main process.
const LOGIN_SHEET_URL =
  'https://docs.google.com/spreadsheets/d/1-cWNTql8wN5Ii7Vx6NZ7Gl74z9yeMzZcZxhjd6JYPm8/edit?usp=sharing'

export interface LoginResult {
  ok: boolean
  /** True when the credentials were correct but this account already has MAX_CONCURRENT_SESSIONS devices signed in (see core/sessionSlots.ts). */
  maxSessions?: boolean
  /** True when the credentials sheet was unreachable (no internet) and this login was allowed through on a locally-remembered credential instead — see electron/offlineAuth.ts. */
  offline?: boolean
}

/**
 * Check a login attempt against the credentials sheet. Fetches fresh every
 * call (no caching) so newly added/changed credentials take effect
 * immediately. Only a plain pass/fail result crosses back to the renderer —
 * never the password, any other column, or the sheet's contents/URL.
 *
 * The login no longer carries any office identity: Zone/Circle/Corporation are
 * chosen by the Head Draughtsman in the app itself (sidebar) and drive document
 * preparation there, so the credentials sheet only needs a Login ID + Password.
 */
export async function validateLogin(loginId: string, password: string): Promise<LoginResult> {
  const table = await importTableFromGoogleLink(LOGIN_SHEET_URL)
  const idHeader = table.headers.find((h) => /login\s*id|user/i.test(h))
  const pwHeader = table.headers.find((h) => /password/i.test(h))
  if (!idHeader || !pwHeader) {
    throw new Error('Could not read the login credentials sheet.')
  }

  const id = loginId.trim().toLowerCase()
  const row = table.rows.find(
    (r) => (r[idHeader] ?? '').trim().toLowerCase() === id && (r[pwHeader] ?? '') === password
  )
  if (!row) return { ok: false }

  return { ok: true }
}

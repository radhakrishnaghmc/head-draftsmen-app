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

/**
 * Login IDs of the other users registered under a corporation, for the Evaluation Sheet's "Share with other
 * users" list. Only the IDs leave this module — never passwords. The sheet is the single source of truth for
 * who belongs to which corporation, so the main process re-checks every share target against it too.
 */
export async function listCorporationUsers(corporation: string, excludeLoginId: string): Promise<string[]> {
  const corp = corporation.trim().toLowerCase()
  const me = excludeLoginId.trim().toLowerCase()
  const ids = new Set<string>()
  for (const u of await loadSheetUsers()) {
    if (u.id.toLowerCase() !== me && u.corporation === corp) ids.add(u.id)
  }
  return [...ids].sort((a, b) => a.localeCompare(b))
}

// The user list is only ever read here (login itself always re-downloads the sheet), and downloading the whole
// sheet on every dialog open was the slow part — so it's kept for a few minutes, with concurrent callers sharing
// one download. Only IDs and corporations are held, never passwords.
const SHEET_USERS_TTL_MS = 10 * 60 * 1000
let sheetUsers: { at: number; users: { id: string; corporation: string }[] } | null = null
let sheetUsersInflight: Promise<{ id: string; corporation: string }[]> | null = null

export function loadSheetUsers(): Promise<{ id: string; corporation: string }[]> {
  if (sheetUsers && Date.now() - sheetUsers.at < SHEET_USERS_TTL_MS) return Promise.resolve(sheetUsers.users)
  if (!sheetUsersInflight) {
    sheetUsersInflight = (async () => {
      const table = await importTableFromGoogleLink(LOGIN_SHEET_URL)
      const idHeader = table.headers.find((h) => /login\s*id|user/i.test(h))
      const corpHeader = table.headers.find((h) => /corporation/i.test(h))
      if (!idHeader || !corpHeader) throw new Error('Could not read the login credentials sheet.')
      const users = table.rows
        .map((r) => ({ id: (r[idHeader] ?? '').trim(), corporation: (r[corpHeader] ?? '').trim().toLowerCase() }))
        .filter((u) => u.id)
      sheetUsers = { at: Date.now(), users }
      return users
    })().finally(() => {
      sheetUsersInflight = null
    })
  }
  return sheetUsersInflight
}

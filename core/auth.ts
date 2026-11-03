// Login results shared between the main process and the renderer. Credentials live in Firestore `accounts`
// (see electron/firebaseSync.ts) — there is no login sheet any more.

export interface LoginResult {
  ok: boolean
  /** Google sign-in only: the Google account has no account here yet — the app asks for the office (and name, mobile) and calls googleSignup. */
  needsOffice?: boolean
  /** The Google account's email, with `needsOffice`. */
  email?: string
  /** True when the credentials were correct but this account already has MAX_CONCURRENT_SESSIONS devices signed in (see core/sessionSlots.ts). */
  maxSessions?: boolean
  /** True when the credentials sheet was unreachable (no internet) and this login was allowed through on a locally-remembered credential instead — see electron/offlineAuth.ts. */
  offline?: boolean
  /** Self-registered accounts only: the offices on the login (the sign-up office first, at most 2). The account is confined to them. */
  offices?: { corporation: string; zone: string; circle: string }[]
  /** The normalised id this login's data is stored under — lets the app tell when a different account takes over a machine. */
  loginId?: string
}

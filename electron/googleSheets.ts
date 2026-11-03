import { app, safeStorage, shell } from 'electron'
import * as fs from 'fs'
import * as path from 'path'
import * as http from 'http'
import * as crypto from 'crypto'
import { GOOGLE_OAUTH_CLIENT_ID, GOOGLE_OAUTH_CLIENT_SECRET } from '../core/googleOAuthConfig'
import { officeTabTitle, gridSignature, chunkRows, columnLetters, quoteTab, worksGridsByOffice } from '../core/onlineWorksSheet'
import type { PersistedState } from '../core/types'
import type { GoogleSheetsStatus } from './ipc-contract'

/**
 * Online Works List: signs the user in with Google (loopback + PKCE) and mirrors each office's Works List into its own
 * tab of one spreadsheet this app creates. The `drive.file` scope only reaches files the app itself created, so it never
 * sees anything else in the user's Drive. One-way: the sheet is never read back.
 */

const SCOPES = 'openid email https://www.googleapis.com/auth/drive.file'
const SHEET_TITLE = 'Head Draughtsman - Online Works List'
const DEBOUNCE_MS = 5000
const ROWS_PER_REQUEST = 2000

interface Account {
  /** Refresh token, encrypted with the OS keychain when available (then `enc` is true). */
  token: string
  enc: boolean
  email: string
  spreadsheetId?: string
  /** Signature of the grid last written per office, so an unchanged Works List isn't re-sent. */
  sigs: Record<string, string>
  lastSync?: number
  lastError?: string
}

const storeFile = () => path.join(app.getPath('userData'), 'google-sheets.json')
let loginId: string | null = null
let accessToken: { value: string; expires: number } | null = null
let timer: NodeJS.Timeout | null = null
let pending: PersistedState | null = null
let chain: Promise<unknown> = Promise.resolve()
let syncing = false

function readStore(): Record<string, Account> {
  try {
    return JSON.parse(fs.readFileSync(storeFile(), 'utf8')) as Record<string, Account>
  } catch {
    return {}
  }
}
function writeStore(s: Record<string, Account>) {
  fs.writeFileSync(storeFile(), JSON.stringify(s), 'utf8')
}
function account(): Account | null {
  return loginId ? readStore()[loginId] ?? null : null
}
function patchAccount(p: Partial<Account>) {
  if (!loginId) return
  const s = readStore()
  if (s[loginId]) writeStore({ ...s, [loginId]: { ...s[loginId], ...p } })
}
function refreshToken(a: Account): string {
  return a.enc ? safeStorage.decryptString(Buffer.from(a.token, 'base64')) : a.token
}
const configured = () => Boolean(GOOGLE_OAUTH_CLIENT_ID && GOOGLE_OAUTH_CLIENT_SECRET)

/** Called on sign-in / sign-out of the app login; a different login has its own Google connection. */
export function setLogin(id: string | null) {
  if (timer) clearTimeout(timer)
  timer = null
  pending = null
  accessToken = null
  loginId = id
}

export function getStatus(): GoogleSheetsStatus {
  const a = account()
  return {
    configured: configured(),
    connected: Boolean(a),
    email: a?.email ?? '',
    spreadsheetUrl: a?.spreadsheetId ? `https://docs.google.com/spreadsheets/d/${a.spreadsheetId}` : '',
    lastSync: a?.lastSync ?? 0,
    lastError: a?.lastError ?? '',
    syncing
  }
}

async function tokenRequest(params: Record<string, string>): Promise<{ access_token: string; expires_in: number; refresh_token?: string; id_token?: string }> {
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: GOOGLE_OAUTH_CLIENT_ID, client_secret: GOOGLE_OAUTH_CLIENT_SECRET, ...params })
  })
  const body = (await res.json().catch(() => ({}))) as { error?: string; error_description?: string }
  if (!res.ok) throw new Error(body.error_description || body.error || `Google sign-in failed (${res.status})`)
  return body as never
}

interface GoogleIdentity {
  email: string
  refresh: string
  access: string
  expires: number
}

/** Stops the sign-in currently waiting on the browser (the user closed the Google tab, so nothing will ever come back). */
let cancelPending: (() => void) | null = null
export function cancelAuthorize(): void {
  cancelPending?.()
}

/** Opens the browser for Google sign-in (loopback + PKCE) and returns the verified account and its tokens. */
async function authorize(): Promise<GoogleIdentity> {
  if (!configured()) throw new Error('Google sign-in is not set up in this build yet.')
  const verifier = crypto.randomBytes(32).toString('base64url')
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url')
  const state = crypto.randomBytes(16).toString('hex')

  const code = await new Promise<{ code: string; redirect: string }>((resolve, reject) => {
    // Known once the server is listening; read before close(), after which the server has no address.
    let redirect = ''
    const server = http.createServer((req, res) => {
      const url = new URL(req.url ?? '/', 'http://127.0.0.1')
      if (url.pathname !== '/') {
        res.writeHead(404).end()
        return
      }
      const err = url.searchParams.get('error')
      const got = url.searchParams.get('code')
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' })
      res.end(`<html><body style="font-family:sans-serif;padding:40px"><h2>${got && url.searchParams.get('state') === state ? 'Signed in. You can close this tab and return to the app.' : 'Sign-in was not completed.'}</h2></body></html>`)
      clearTimeout(to)
      cancelPending = null
      server.close()
      if (err || !got || url.searchParams.get('state') !== state) reject(new Error(err ? `Google: ${err}` : 'Sign-in was cancelled.'))
      else resolve({ code: got, redirect })
    })
    const to = setTimeout(() => {
      cancelPending = null
      server.close()
      reject(new Error('Sign-in timed out. Try again.'))
    }, 180000)
    cancelPending = () => {
      cancelPending = null
      clearTimeout(to)
      server.close()
      reject(new Error('Sign-in was cancelled.'))
    }
    server.listen(0, '127.0.0.1', () => {
      redirect = `http://127.0.0.1:${(server.address() as { port: number }).port}`
      const q = new URLSearchParams({
        client_id: GOOGLE_OAUTH_CLIENT_ID,
        redirect_uri: redirect,
        response_type: 'code',
        scope: SCOPES,
        code_challenge: challenge,
        code_challenge_method: 'S256',
        access_type: 'offline',
        prompt: 'consent',
        state
      })
      void shell.openExternal(`https://accounts.google.com/o/oauth2/v2/auth?${q}`)
    })
  })

  const tok = await tokenRequest({ code: code.code, code_verifier: verifier, redirect_uri: code.redirect, grant_type: 'authorization_code' })
  if (!tok.refresh_token) throw new Error('Google did not return a refresh token. Remove the app from your Google account permissions and try again.')
  const claims = tok.id_token ? (JSON.parse(Buffer.from(tok.id_token.split('.')[1], 'base64url').toString()) as { email?: string; email_verified?: boolean }) : {}
  if (!claims.email || claims.email_verified === false) throw new Error('Google did not confirm an email address for this account.')
  return { email: claims.email.toLowerCase(), refresh: tok.refresh_token, access: tok.access_token, expires: Date.now() + (tok.expires_in - 60) * 1000 }
}

function storeIdentity(id: GoogleIdentity) {
  if (!loginId) return
  const enc = safeStorage.isEncryptionAvailable()
  const stored = enc ? safeStorage.encryptString(id.refresh).toString('base64') : id.refresh
  const prev = readStore()[loginId]
  // Reconnecting the same Google account keeps its spreadsheet; a different account starts fresh.
  const keep = prev && prev.email === id.email ? { spreadsheetId: prev.spreadsheetId, sigs: prev.sigs } : { sigs: {} as Record<string, string> }
  writeStore({ ...readStore(), [loginId]: { token: stored, enc, email: id.email, ...keep } })
  accessToken = { value: id.access, expires: id.expires }
}

/** Settings → Online Works List: connects a Google account to the signed-in login. */
export async function signIn(): Promise<GoogleSheetsStatus> {
  if (!loginId) throw new Error('Sign in to the app first.')
  storeIdentity(await authorize())
  return getStatus()
}

/** Login screen: verifies the Google account (the caller decides whether it is allowed in). The tokens are kept by `adopt`. */
export async function authorizeForLogin(): Promise<{ email: string; adopt: () => void }> {
  const id = await authorize()
  return { email: id.email, adopt: () => storeIdentity(id) }
}

export async function signOut(): Promise<GoogleSheetsStatus> {
  const a = account()
  if (a && loginId) {
    try {
      await fetch(`https://oauth2.googleapis.com/revoke?token=${encodeURIComponent(refreshToken(a))}`, { method: 'POST' })
    } catch {
      /* revoking is best-effort; the local credentials are removed regardless */
    }
    const s = readStore()
    delete s[loginId]
    writeStore(s)
  }
  accessToken = null
  if (timer) clearTimeout(timer)
  timer = null
  pending = null
  return getStatus()
}

async function getAccessToken(): Promise<string> {
  if (accessToken && accessToken.expires > Date.now()) return accessToken.value
  const a = account()
  if (!a) throw new Error('Not signed in to Google.')
  const tok = await tokenRequest({ refresh_token: refreshToken(a), grant_type: 'refresh_token' })
  accessToken = { value: tok.access_token, expires: Date.now() + (tok.expires_in - 60) * 1000 }
  return accessToken.value
}

async function sheetsApi<T>(method: string, url: string, body?: unknown): Promise<T> {
  for (let attempt = 0; ; attempt++) {
    const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets${url}`, {
      method,
      headers: { Authorization: `Bearer ${await getAccessToken()}`, 'Content-Type': 'application/json' },
      body: body === undefined ? undefined : JSON.stringify(body)
    })
    // Per-minute quota: wait out the window rather than failing the whole push.
    if (res.status === 429 && attempt < 4) {
      await new Promise((r) => setTimeout(r, 15000 * (attempt + 1)))
      continue
    }
    const json = (await res.json().catch(() => ({}))) as { error?: { message?: string; status?: string } }
    if (!res.ok) {
      const err = new Error(json.error?.message ?? `Google Sheets error ${res.status}`) as Error & { status: number }
      err.status = res.status
      throw err
    }
    return json as T
  }
}

interface SheetMeta {
  properties: { sheetId: number; title: string }
}

async function ensureSpreadsheet(): Promise<string> {
  const a = account()
  if (a?.spreadsheetId) {
    try {
      await sheetsApi('GET', `/${a.spreadsheetId}?fields=spreadsheetId`)
      return a.spreadsheetId
    } catch (e) {
      // Deleted by the user (or trashed): make a fresh one and re-send everything.
      if (![403, 404].includes((e as { status?: number }).status ?? 0)) throw e
      patchAccount({ spreadsheetId: undefined, sigs: {} })
    }
  }
  const created = await sheetsApi<{ spreadsheetId: string }>('POST', '', { properties: { title: SHEET_TITLE } })
  patchAccount({ spreadsheetId: created.spreadsheetId })
  return created.spreadsheetId
}

async function writeTab(id: string, title: string, grid: string[][]) {
  const meta = await sheetsApi<{ sheets: SheetMeta[] }>('GET', `/${id}?fields=sheets.properties(sheetId,title)`)
  let tab = meta.sheets.find((s) => s.properties.title === title)
  const rows = Math.max(grid.length, 1)
  const cols = Math.max(...grid.map((r) => r.length), 1)
  const props = { title, gridProperties: { rowCount: rows + 20, columnCount: cols, frozenRowCount: 1 } }
  if (!tab) {
    const r = await sheetsApi<{ replies: { addSheet: SheetMeta }[] }>('POST', `/${id}:batchUpdate`, { requests: [{ addSheet: { properties: props } }] })
    tab = r.replies[0].addSheet
  } else {
    await sheetsApi('POST', `/${id}:batchUpdate`, {
      requests: [{ updateSheetProperties: { properties: { sheetId: tab.properties.sheetId, gridProperties: props.gridProperties }, fields: 'gridProperties(rowCount,columnCount,frozenRowCount)' } }]
    })
  }
  const q = quoteTab(title)
  await sheetsApi('POST', `/${id}/values/${encodeURIComponent(q)}:clear`, {})
  let start = 1
  for (const part of chunkRows(grid, ROWS_PER_REQUEST)) {
    const range = `${q}!A${start}:${columnLetters(cols)}${start + part.length - 1}`
    await sheetsApi('PUT', `/${id}/values/${encodeURIComponent(range)}?valueInputOption=RAW`, { range, majorDimension: 'ROWS', values: part })
    start += part.length
  }
  // The new spreadsheet starts with an empty "Sheet1"; drop it once a real tab exists.
  const dflt = meta.sheets.find((s) => s.properties.title === 'Sheet1')
  if (dflt && dflt.properties.title !== title) {
    await sheetsApi('POST', `/${id}:batchUpdate`, { requests: [{ deleteSheet: { sheetId: dflt.properties.sheetId } }] }).catch(() => undefined)
  }
}

async function pushNow(state: PersistedState, force: boolean) {
  if (!account()) return
  const grids = worksGridsByOffice(state)
  const sigs = { ...(account()?.sigs ?? {}) }
  const changed = Object.entries(grids).filter(([k, g]) => force || sigs[k] !== gridSignature(g))
  if (changed.length === 0) return
  syncing = true
  try {
    const id = await ensureSpreadsheet()
    // ensureSpreadsheet may have reset the signatures for a recreated sheet.
    const fresh = account()?.sigs ?? {}
    for (const [key, grid] of changed) {
      if (!force && fresh[key] === gridSignature(grid)) continue
      await writeTab(id, officeTabTitle(key), grid)
      fresh[key] = gridSignature(grid)
      patchAccount({ sigs: fresh })
    }
    patchAccount({ lastSync: Date.now(), lastError: '' })
  } catch (e) {
    patchAccount({ lastError: (e as Error).message })
    throw e
  } finally {
    syncing = false
  }
}

/** Called after every workspace save: pushes only the offices whose Works List changed, a few seconds after the last edit. */
export function scheduleSync(state: PersistedState) {
  if (!account()) return
  pending = state
  if (timer) clearTimeout(timer)
  timer = setTimeout(() => {
    const s = pending
    pending = null
    if (s) chain = chain.catch(() => undefined).then(() => pushNow(s, false)).catch(() => undefined)
  }, DEBOUNCE_MS)
}

/** Settings → "Update online sheet now": re-sends every office's Works List. Resolves the reason it failed, or ''. */
export async function syncNow(state: PersistedState): Promise<string> {
  if (!account()) return 'Sign in with Google first.'
  try {
    const run = chain.catch(() => undefined).then(() => pushNow(state, true))
    chain = run
    await run
    return ''
  } catch (e) {
    return (e as Error).message
  }
}

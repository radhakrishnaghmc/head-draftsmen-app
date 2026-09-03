// Reads the Telangana Public Health department's "Cement & Steel Rates"
// download page (Downloads > Cement & Steel Rates, oId=601 under the
// Downloads menu id=221 — https://publichealth.telangana.gov.in/getInfo.do?
// dt=2&id=221&oId=601). That page has no static list: a hidden sub-form
// posts back to /getInfo.do with mode=getCementSteelRates once a session +
// CSRF token are established, and the department's own template only
// renders the year/month <select> options (and the actual row table) after
// that POST — so this walks the same three-request flow a browser does:
//   1) GET  /home.do                          — obtains the session cookie
//   2) GET  /getInfo.do?dt=2&id=221&oId=601   — obtains a page csrfToken
//   3) POST /getInfo.do  mode=getCementSteelRates&reportYear=all&quarter=all
// Each row's "click here for more details" link carries a short-lived
// (5 hour) signed JWT whose `sub` claim is the file's path on the server;
// the page hands that token straight to /download.do?fileName=<token>,
// which needs no session of its own — the token is self-contained.

import type { CementSteelRate } from '../core/cementSteelRates'

const BASE = 'https://publichealth.telangana.gov.in'
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) HeadDraftsmenApp'

class CookieJar {
  private cookies = new Map<string, string>()
  header(): string {
    return Array.from(this.cookies.entries())
      .map(([k, v]) => `${k}=${v}`)
      .join('; ')
  }
  absorb(res: Response) {
    const setCookies = res.headers.getSetCookie?.() ?? []
    for (const c of setCookies) {
      const pair = c.split(';')[0]
      const eq = pair.indexOf('=')
      if (eq > 0) this.cookies.set(pair.slice(0, eq).trim(), pair.slice(eq + 1).trim())
    }
  }
}

/** Best-effort file extension from the token's own (unverified) `sub` claim — cosmetic only, never trusted for auth. */
function extFromToken(token: string): string {
  try {
    const payload = token.split('.')[1]
    const json = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8')) as { sub?: string }
    const m = /\.([a-zA-Z0-9]{2,5})$/.exec(json.sub ?? '')
    return m ? m[1].toLowerCase() : 'pdf'
  } catch {
    return 'pdf'
  }
}

export async function fetchCementSteelRates(): Promise<CementSteelRate[]> {
  const jar = new CookieJar()

  const homeRes = await fetch(`${BASE}/home.do`, { headers: { 'User-Agent': USER_AGENT } })
  jar.absorb(homeRes)
  await homeRes.text()

  const infoUrl = `${BASE}/getInfo.do?dt=2&id=221&oId=601`
  const infoRes = await fetch(infoUrl, {
    headers: { 'User-Agent': USER_AGENT, Referer: `${BASE}/home.do`, Cookie: jar.header() }
  })
  jar.absorb(infoRes)
  const infoHtml = await infoRes.text()
  const csrfToken = /name="csrfToken"\s+value="([^"]+)"/.exec(infoHtml)?.[1]
  if (!csrfToken) {
    throw new Error("Couldn't open the Cement & Steel Rates page — the department site may be down or its layout has changed.")
  }

  const body = new URLSearchParams({ csrfToken, mode: 'getCementSteelRates', reportYear: 'all', quarter: 'all' })
  const listRes = await fetch(`${BASE}/getInfo.do`, {
    method: 'POST',
    headers: {
      'User-Agent': USER_AGENT,
      'Content-Type': 'application/x-www-form-urlencoded',
      Referer: infoUrl,
      Cookie: jar.header()
    },
    body: body.toString()
  })
  const html = await listRes.text()

  // Each row: Sl.No cell, description cell (icon + "click here for more
  // details" link carrying the download token), File Posted date cell.
  const rowRe =
    /<td class="tBody" align="center">\s*(\d+)\s*<\/td>[\s\S]*?<td class="tBody">([\s\S]*?)<\/td>[\s\S]*?<td class="tBody">\s*([\d/]+)\s*<\/td>/g

  const out: CementSteelRate[] = []
  let m: RegExpExecArray | null
  while ((m = rowRe.exec(html))) {
    const [, slNoStr, descHtml, datePosted] = m
    const token = /getDownload\('([^']+)'\)/.exec(descHtml)?.[1]
    if (!token) continue
    let description = descHtml
      .split('<a ')[0]
      .replace(/<[^>]+>/g, ' ')
      .replace(/\.{3,}/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim()
    // A recurring typo on the department's own template drops the leading
    // "C" off the very latest entry ("ement & Steel Rates …") — cosmetic fix.
    if (description.startsWith('ement ')) description = `C${description}`
    out.push({ slNo: Number(slNoStr), description, datePosted, token, ext: extFromToken(token) })
  }

  if (out.length === 0) {
    throw new Error('No Cement & Steel Rate circulars were found on the department site.')
  }
  return out
}

export async function downloadCementSteelRateBuffer(token: string): Promise<Buffer> {
  const res = await fetch(`${BASE}/download.do?fileName=${encodeURIComponent(token)}&filePath=no`, {
    headers: { 'User-Agent': USER_AGENT }
  })
  if (!res.ok) throw new Error('The department site refused the download — the link may have expired, try again.')
  return Buffer.from(await res.arrayBuffer())
}

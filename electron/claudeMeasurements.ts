import { execFile } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { SEIG_CHART } from '../core/seigniorageChart'

const chartList = SEIG_CHART.map((c) => `${c.sl}=${c.name}`).join('; ')

const pagePrompt = (page: number) =>
  `Read page ${page} only of the PDF "input.pdf" in the current directory (Read tool, pages: "${page}"). It is a page of a handwritten Indian PWD/GHMC measurement book. ` +
  'Reply with ONLY one JSON object, no prose, no code fence: ' +
  '{"header":{"work":"","agency":"","estimate":"","contract":"","percentage":"","markOut":"","completion":""},' +
  '"items":[{"continued":false,"description":"item description","unit":"Cum|Sqm|Rmt|Nos","lines":[{"location":"","no":1,"l":0,"b":[0,0,0],"d":[0,0,0],"written":0}],"totalWritten":0}],' +
  '"abstract":[{"description":"","qtyWritten":0,"unit":"","rate":0,"amountWritten":0,"chartSl":null}]}. ' +
  'Rules: "header" only from a cover/title page, else empty strings. "items" are measurement entries: one object per work item on this page; ' +
  'set "continued":true if the page continues the previous page\'s item (no new description) — then description may be empty. ' +
  'Each line: no. of units, length l, up to three breadths in b and three depths/heights in d (null when absent), and "written" = the contents written for that line. ' +
  '"abstract" is only for an abstract-of-cost page (item, quantity, unit, rate, amount); else []. ' +
  `For every abstract item set "chartSl" to the number of the matching seigniorage chart item, or null if none matches: ${chartList}. ` +
  'Use null for anything unreadable; never guess numbers. Empty page → {"header":{},"items":[],"abstract":[]}.'

// The one `claude` process in flight (login or a page read), so Stop can kill it.
let activeChild: ReturnType<typeof execFile> | null = null
export const cancelClaude = (): void => void activeChild?.kill()

// Drives the user's own `claude` CLI — their account, no API key or token handled here.
function runClaude(args: string[], cwd: string | undefined, timeout: number): Promise<string> {
  // Packaged apps launch with a minimal PATH; add the usual CLI install dirs.
  const PATH = `${process.env.PATH ?? ''}:/opt/homebrew/bin:/usr/local/bin:${process.env.HOME}/.local/bin`
  return new Promise((resolve, reject) => {
    activeChild = execFile('claude', args, { cwd, env: { ...process.env, PATH }, timeout, maxBuffer: 20 * 1024 * 1024 }, (err, stdout, stderr) =>
      err ? reject(new Error(`Claude CLI failed (is it installed?): ${stderr || err.message}`)) : resolve(stdout)
    )
    activeChild.on('exit', () => (activeChild = null))
  })
}

export async function claudeAuthStatus(): Promise<boolean> {
  try {
    return JSON.parse(await runClaude(['auth', 'status'], undefined, 30_000)).loggedIn === true
  } catch {
    return false
  }
}

// Opens the browser sign-in; resolves when the user finishes (or rejects on timeout/failure).
export async function claudeLogin(): Promise<void> {
  await runClaude(['auth', 'login'], undefined, 5 * 60_000)
  if (!(await claudeAuthStatus())) throw new Error('Claude sign-in was not completed.')
}

export async function extractMeasurementsWithClaude(pdfBytes: Uint8Array, page: number): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'hda-claude-'))
  try {
    await writeFile(join(dir, 'input.pdf'), pdfBytes)
    return await runClaude(['-p', pagePrompt(Math.max(1, Math.floor(page))), '--allowedTools', 'Read', '--output-format', 'text'], dir, 5 * 60_000)
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

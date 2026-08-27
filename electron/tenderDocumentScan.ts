// Recursively collects L1 sheets / Online Intimations out of whatever mix of
// loose files and folders the user picked in pickTenderDocuments' dialog —
// so "Update from L1/LOA" and "Address from Intimation" can process an
// entire office's Tender Evaluations tree in one click instead of hunting
// down each work's own PDF one at a time.
//
// Every fs call here is async (fs.promises), not fs.*Sync — this runs on the
// Electron MAIN process, which is single-threaded and shared with every
// window's IPC and the UI itself; a synchronous walk over a folder holding
// hundreds/thousands of files (readdirSync + statSync per entry, then
// readFileSync of every match) froze the ENTIRE app for the whole scan, not
// just this dialog. The async versions hand each disk op off to libuv's
// thread pool, so the main process's event loop stays free between them.
import * as fs from 'fs'
import * as path from 'path'
import type { PickedTenderDocument } from './ipc-contract'

const fsp = fs.promises

// Filenames worth reading, once INSIDE a folder the user picked — an L1
// selection sheet / Commercial Evaluation page, or an Online Intimation — as
// opposed to the hundreds of other PDFs (PAN cards, GST certs, ITRs,
// litigation declarations, turnover certificates, …) that sit alongside them
// in a real office's folder tree (see resources/… — a real "Common
// Documents" folder for one bidder alone routinely holds a dozen of these).
// Matches this app's own generated filenames too ("Intimation - X.pdf").
// Deliberately permissive — a false positive just gets read and silently
// skipped downstream by parseTenderEvaluation/parseIntimationNotice finding
// nothing useful in it — the goal is avoiding thousands of needless KYC-
// document reads, not perfect recall. Only applied to files discovered by
// recursing into a folder; a file the user explicitly picked by name is
// always included regardless of what it's called.
// "eval" alone (not just "commercial eval") — a real file found in
// Nizampet-58's own folder tree is simply named "evaluation sheet.pdf".
const RELEVANT_NAME = /l-?1|stage\s*selected|eval|intimation|loa\b/i

const RELEVANT_EXT = new Set(['.pdf', '.html', '.htm'])

// A folder like a full circle's "Tender Evaluations" can hold thousands of
// files once every bidder's KYC documents are counted — this caps how many
// this app will actually read back into memory and hand to the renderer for
// parsing in one go, rather than trying to process an accidentally-too-broad
// folder pick (e.g. the user's whole Google Drive root) wholesale.
const MAX_FILES = 1000

// A local read is milliseconds; this is generous enough to cover a slow
// network drive or a cloud-sync client that has to fetch a placeholder file
// on demand, while still bounding the worst case to something a user won't
// mistake for the app having frozen.
const FILE_READ_TIMEOUT_MS = 20_000

/** `done` files matched so far during the recursive directory walk (no fixed total — folder sizes aren't known upfront) or files actually read back into memory (fixed `total`, so this phase can show a real percentage). */
export type TenderScanProgress = (phase: 'scanning' | 'reading', done: number, total: number) => void

async function walk(dir: string, out: string[], onProgress?: TenderScanProgress): Promise<void> {
  if (out.length >= MAX_FILES) return
  let entries: fs.Dirent[]
  try {
    entries = await fsp.readdir(dir, { withFileTypes: true })
  } catch {
    return // permission denied, broken symlink, …
  }
  for (const entry of entries) {
    if (out.length >= MAX_FILES) return
    const full = path.join(dir, entry.name)
    if (entry.isDirectory()) {
      await walk(full, out, onProgress)
    } else if (RELEVANT_EXT.has(path.extname(full).toLowerCase()) && RELEVANT_NAME.test(entry.name)) {
      out.push(full)
      onProgress?.('scanning', out.length, 0)
    }
  }
}

export async function collectTenderDocuments(
  pickedPaths: string[],
  onProgress?: TenderScanProgress
): Promise<PickedTenderDocument[]> {
  const files: string[] = []
  for (const p of pickedPaths) {
    let stat: fs.Stats
    try {
      stat = await fsp.stat(p)
    } catch {
      continue
    }
    if (stat.isDirectory()) {
      await walk(p, files, onProgress)
    } else {
      // Explicitly picked by name — always included, no filename filter.
      files.push(p)
    }
  }

  const capped = files.slice(0, MAX_FILES)
  const results: PickedTenderDocument[] = []
  for (let i = 0; i < capped.length; i++) {
    const p = capped[i]
    try {
      // A single stalled read (a OneDrive/Google Drive/Dropbox "on-demand"
      // placeholder that has to download first, or a network drive that's
      // dropped) would otherwise block this whole sequential loop forever —
      // real report: the scan reads every file it finds, then just sits on
      // "Reading N/total" with no error, no timeout of its own. Skip past a
      // file that doesn't resolve in time instead of hanging the entire
      // update. fsp.readFile itself isn't cancelled by the race losing (Node
      // has no clean way to abort an in-flight fs read) — it's just left to
      // finish or fail on its own and its result is ignored either way.
      const bytes = await Promise.race([
        fsp.readFile(p),
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error('timed out')), FILE_READ_TIMEOUT_MS)
        )
      ])
      results.push({ name: path.basename(p), bytes: new Uint8Array(bytes) })
    } catch {
      // Unreadable or timed out — skip it and keep the rest of the scan moving.
    }
    onProgress?.('reading', i + 1, capped.length)
  }
  return results
}

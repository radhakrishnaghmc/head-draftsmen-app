// Recursively collects L1 sheets / Online Intimations out of whatever mix of
// loose files and folders the user picked in pickTenderDocuments' dialog —
// so "Update from L1/LOA" and "Address from Intimation" can process an
// entire office's Tender Evaluations tree in one click instead of hunting
// down each work's own PDF one at a time.
import * as fs from 'fs'
import * as path from 'path'
import type { PickedTenderDocument } from './ipc-contract'

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

function walk(dir: string, out: string[]): void {
  if (out.length >= MAX_FILES) return
  let entries: string[]
  try {
    entries = fs.readdirSync(dir)
  } catch {
    return // permission denied, broken symlink, …
  }
  for (const entry of entries) {
    if (out.length >= MAX_FILES) return
    const full = path.join(dir, entry)
    let stat: fs.Stats
    try {
      stat = fs.statSync(full)
    } catch {
      continue
    }
    if (stat.isDirectory()) {
      walk(full, out)
    } else if (RELEVANT_EXT.has(path.extname(full).toLowerCase()) && RELEVANT_NAME.test(entry)) {
      out.push(full)
    }
  }
}

export function collectTenderDocuments(pickedPaths: string[]): PickedTenderDocument[] {
  const files: string[] = []
  for (const p of pickedPaths) {
    let stat: fs.Stats
    try {
      stat = fs.statSync(p)
    } catch {
      continue
    }
    if (stat.isDirectory()) {
      walk(p, files)
    } else {
      // Explicitly picked by name — always included, no filename filter.
      files.push(p)
    }
  }
  return files.slice(0, MAX_FILES).map((p) => ({ name: path.basename(p), bytes: new Uint8Array(fs.readFileSync(p)) }))
}

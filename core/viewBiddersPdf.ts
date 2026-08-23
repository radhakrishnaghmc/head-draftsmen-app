// Parses the Telangana e-procurement portal's "View Bidders" / "Evaluation
// Stages" page (its "Supplier List / Commercial Stage" screen, saved as PDF):
// the list of contractors who participated in a tender. The tender header
// (Tender ID, NIT No, Name of Work, ECV, bid dates) shares the exact layout of
// the Responsiveness / Commercial-Evaluation pages, so that is read with the
// existing parseTenderEvaluation (core/tenderEvaluationPdf.ts, itself a thin
// composer over core/tenderAgents/ — see that folder for the actual
// per-field detectors); this module only adds the supplier-row reader, which
// the Evaluation Sheet needs to lay out one column per bidder.

import { isNameContinuation, prevNonEmptyLine } from './tenderAgents/shared'

/** A participating bidder from the supplier list (name is the only field the
 * Evaluation Sheet uses; the rest are captured for display / future use). */
export interface ParticipatingBidder {
  name: string
  /** Supplier Bid Number, e.g. "1340371-0". */
  bidNumber: string
  /** "04/08/2026 02:19 PM" — Bid Submitted Date & Time. */
  submittedAt: string
}

// The numeric core of a supplier row: "<BidNo like 1340371-0> <dd/mm/yyyy hh:mm
// AM/PM> <IPv4>". The company name sits on the same line just before it, but a
// long name wraps onto its own line(s) with the numbers landing on a line
// *between* the wrapped name halves (the numbers are vertically centred against
// the taller name cell) — e.g. "BOBBA RAVI CHANDRA CIVIL" / "1339995-1 …
// 106.222.233.188" / "CONTRACTOR". So when the number row carries no name
// prefix, take the name from the line above plus the line below when that's a
// name continuation.
const ROW_CORE = /(\d{5,8}-\d)\s+(\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}\s*[AP]M)\s+\d{1,3}(?:\.\d{1,3}){3}/i

/**
 * Every participating bidder from the "Supplier List" table, in the page's
 * order. Returns [] when the page carries no supplier rows (e.g. a different
 * portal screen). Names are best-effort against this page's specific layout.
 */
export function parseParticipatingBidders(lines: string[]): ParticipatingBidder[] {
  const out: ParticipatingBidder[] = []
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim()
    const m = ROW_CORE.exec(line)
    if (!m) continue

    let name = line.slice(0, m.index).trim()
    if (!name) {
      const prev = prevNonEmptyLine(lines, i)
      const next = lines[i + 1]?.trim() ?? ''
      name = isNameContinuation(next) ? `${prev} ${next}`.trim() : prev
    }
    name = name.replace(/\s+/g, ' ').trim()
    if (name) out.push({ name, bidNumber: m[1], submittedAt: m[2].replace(/\s+/g, ' ').trim() })
  }
  return out
}

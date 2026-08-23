import { joinLines } from './shared'

/**
 * Agent: Tender ID — the portal's own internal numeric ID for this tender
 * (e.g. "726879"), read from the L1/Commercial-Evaluation sheet's header.
 * Independent of every other agent: reads straight off the reconstructed
 * page lines, does not depend on NIT No, Name of Work, or anything else
 * having been found first.
 */
export function detectTenderId(lines: string[]): string | undefined {
  const joined = joinLines(lines)

  // Sits in the header's right column; pdf.js emits it either as
  // "Tender ID <id>" or — when the value cell rounds a hair above its label —
  // as "<id> Tender ID" (value before label). Accept both orders. The id is a
  // 6-digit number, so require 5–8 digits: on some sheets (e.g. the SE office
  // layout) the NIT date's year — "…Dt. 27.07.2026 Tender ID 722264…" — sits
  // right before the label, and an unconstrained "\d+" matched that "2026"
  // (leftmost) instead of the real id.
  // A third layout runs the two columns' values together on one row with
  // both labels on a later row — "720716 13/DB/EE/…/2026-27 Tender ID
  // Notice Number …" — so the id ends up adjacent to the NIT No's own
  // opening ("<id> <nitNo>"), nowhere near the "Tender ID" label at all. That
  // opening is usually a fresh "/DB/EE/" continuation, but when the NIT No's
  // own label already consumed "…/DB/EE/…Circle-" earlier in the text, only
  // the wrapped TAIL follows the id instead — "…Circle- Enquiry/IFB/Tender
  // 699966 58/QBZ/CMC/2026-27 Dt.08.05.2026…" — no "/DB/EE/" in that tail at
  // all, just "<circleNo>/[QBZ/]CMC/<year>" directly. Accept either shape.
  const m =
    /Tender ID\s+(\d{5,8})\b|\b(\d{5,8})\s+Tender ID|\b(\d{5,8})\s+(?=\d{1,3}\/(?:DB\/EE\/|(?:QBZ\/)?CMC\/))/i.exec(
      joined
    )
  return m ? (m[1] ?? m[2] ?? m[3]) : undefined
}

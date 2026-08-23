import { joinLines } from './shared'

export interface NitNoAndDate {
  /** NIT / Notice number (the trailing "item N Dated:…" tail is dropped). */
  noticeNo?: string
  /** The "Dated:"/"Dt:" date carried in the NIT No line (e.g. "15.07.2026"), if present. */
  noticeDate?: string
}

// Some offices prefix the item number with 1-2 short alphanumeric segments
// ("E1/06/23/DB/EE/…" instead of a plain "23/DB/EE/…") — each segment is at
// most 2 letters plus up to 3 digits, always carrying a digit, so it can never
// match a label word like "Tender"/"Enquiry"/"IFB" (all-letters) or a 5-8
// digit Tender ID (too many digits for one segment) wedged in front of the
// real code.
const NIT_CODE_PREFIX = '(?:[A-Za-z]{0,2}\\d{1,3}\\/){0,2}'

// Tightens OCR/layout spacing inside an identifier so a code split across the
// PDF's two-column layout ("… Circle- 58/CMC …") rejoins ("…Circle-58/CMC…").
function tightenCode(s: string): string {
  return s
    .replace(/\s*([/\-])\s*/g, '$1')
    .replace(/\s+/g, ' ')
    .trim()
}

/**
 * Reassemble a NIT No that pdf.js interleaved with the header's right column.
 * These NITs read "<code>/DB/EE/<place> Circle-<circleNo>/[QBZ/]CMC/<year>" (a
 * named-place office) or "<code>/DB/EE-<div>/<dept>/C-<circleNo>/[QBZ/]CMC/<year>"
 * (a division office using the short "C-<n>" circle code instead), but the
 * Tender ID, an "Enquiry/IFB/Tender Notice" label fragment, or a stray digit run
 * routinely lands between the wrapped "…Circle-"/"…C-" prefix and the
 * "<circleNo>/[QBZ/]CMC/<year>" tail (e.g. "…Circle- Enquiry/IFB/Tender 699588
 * 57/QBZ/CMC/2026-27", or "…/C- Enquiry/IFB/Tender 723483 54/QBZ/CMC/2026-27").
 * Keep just the prefix and that tail, dropping whatever's wedged between — and,
 * because the tail ends at the year, this also trims any trailing "date:…" /
 * work-name text the value regex over-captured. Falls back to the tightened raw
 * value when the string doesn't fit either canonical shape.
 */
function cleanNit(raw: string): string {
  // Drop any label preamble before the NIT's own code — e.g. "NIT No" landing
  // right next to a secondary "Enquiry/IFB/Tender <tenderId> " sub-label on
  // some L1 layouts leaves that whole label wedged in front of the real
  // "13/DB/EE/…" code ("Enquiry/IFB/Tender 720716 13/DB/EE/…") — but a real
  // "E1/06/" item-number prefix right before "/DB/EE/" is kept (NIT_CODE_PREFIX).
  const codeStart = new RegExp(`${NIT_CODE_PREFIX}\\d{1,3}\\/DB\\/EE\\b`, 'i').exec(raw)
  const s = codeStart ? raw.slice(codeStart.index) : raw
  // "Circle-?": some sheets run the place name straight into the circle
  // number with no hyphen at all ("NizampetCircle58"), not just "Circle-58".
  const m = /^(.*?(?:Circle-?|\/C-))\s*(?:.*?\s)?(\d{1,3}\/(?:QBZ\/)?CMC\/\d{4}\s*-\s*\d{2,4})/i.exec(s)
  return tightenCode(m ? `${m[1]}${m[2]}` : s)
}

/**
 * Agent: NIT No + NIT Date — the tender notice number (e.g.
 * "17/DB/EE/NizampetCircle58/CMC/2026-27") and the date printed right after
 * it ("ITEM 4 Dated:18.08.2026"), read from the L1 sheet's header. Combined
 * into one detector (not two) because both come off the exact same anchor —
 * the NIT No line — so splitting them would mean re-finding that anchor
 * twice and risking the two answers disagreeing about which line it's on.
 */
export function detectNitNoAndDate(lines: string[]): NitNoAndDate {
  const result: NitNoAndDate = {}
  const joined = joinLines(lines)

  // NIT No spans the header's left column and is interrupted by the right
  // column ("Tender ID <id>" / the "Notice Number" label) landing between its
  // wrapped halves in reading order — e.g. "…Circle- 699549 Tender ID
  // 57/QBZ/…". Strip that Tender ID label+value in either order and the
  // "Notice Number" label so the NIT's two halves rejoin, then take what's
  // between "NIT No." and the "item"/"Dt"/"Dated" tail. Same 5–8 digit
  // constraint as the Tender ID agent's own regex — an unconstrained
  // "\d+ Tender ID" ate the trailing "27" off a wrapped "…CMC/2026-27 Tender
  // ID…" (a bare "Tender ID" label with no value of its own on that row),
  // truncating the NIT No to "…CMC/2026-".
  const cleaned = joined
    .replace(/Tender ID\s+\d{5,8}\b|\b\d{5,8}\s+Tender ID/gi, ' ')
    .replace(/Notice Number/gi, ' ')
  // Capture through to the tail (or the line's end), then cleanNit drops any
  // Tender ID / "Enquiry/IFB/Tender" / date fragment wedged into the value.
  const nit = /NIT No\.?\s*(.+?)\s*(?:item\b|Dated\b|Dt\b|Name of Work\b|$)/i.exec(cleaned)
  if (nit) {
    const value = cleanNit(nit[1])
    if (value) result.noticeNo = value
  }
  // Many L1 sheets label this field only "Enquiry/IFB/Tender … Notice Number"
  // with no "NIT No." prefix at all (e.g. "Enquiry/IFB/Tender 08/DB/EE/Nizampet
  // Circle-58/CMC/2026- … Notice Number 27(Item No.01),Dt:25.06.2026"), so the
  // label-anchored match above finds nothing. Fall back to the notice number's
  // own canonical shape wherever it sits in the (label-stripped) text —
  // "<code>/DB/EE/<place> Circle-<circleNo>/[QBZ/]CMC/<year>" — tolerating the
  // spaces the stripped "Tender ID"/"Notice Number" labels leave inside the
  // wrapped "…/2026-  27" tail. tightenCode then rejoins it.
  //
  // "/EE/" and "/CMC/" are both made optional: real Nizampet Circle-58
  // sheets for older/completed works print shorter codes that skip one or
  // both — "E1/06/01/DB/EE/Nizampetcircle-58/2026-27" (no CMC at all) and
  // "E1/06/02/DB/Nizampet Circle-58/2026-27" (no EE either). A regex that
  // requires both literally just returns undefined for these — a real gap,
  // not a rare one, on this same office's own older tenders.
  if (!result.noticeNo) {
    const canonical = new RegExp(
      `${NIT_CODE_PREFIX}\\d{1,3}\\/DB\\/(?:EE\\/)?.+?Circle\\s*-\\s*\\d{1,3}\\/(?:(?:QBZ\\/)?CMC\\/)?\\d{4}\\s*-\\s*\\d{2,4}`,
      'i'
    ).exec(cleaned)
    if (canonical) result.noticeNo = tightenCode(canonical[0])
  }

  // A different office (Serilingampally/Ameenpur, "ee-ptcu-ghmc") prints a
  // structurally different code with no "Circle-<n>"/"/C-<n>" marker at all —
  // "Engg-21/CMC/AMPR-C-47/2026-27". Tried only as a last resort (the DB/EE
  // canonical above already covers every office seen so far and stays
  // untouched) — kept narrow to the one real sample seen from this office
  // rather than guessing at variants that haven't actually been observed.
  if (!result.noticeNo) {
    const engg = /Engg-\d+\/(?:[A-Za-z]+\/)?[A-Za-z0-9-]+\/\d{4}\s*-\s*\d{2,4}/i.exec(cleaned)
    if (engg) result.noticeNo = tightenCode(engg[0])
  }

  // The NIT line carries the notice's own date as "Dated:15.07.2026" /
  // "Dt: 15-07-2026" — split out as the Tender notice Date. "Dated"/"Dt" is
  // the reliable anchor (bid-submission/server dates on the page carry no
  // such label), so an unrelated date is never misread as the notice date.
  // No \b before the anchor: the item number often runs straight into it with
  // no space ("ITEM 7Dated:24.07.2026"), which a word boundary would reject
  // since digit-into-letter isn't a boundary — the lookbehind instead just
  // rules out landing inside another word ("Updated"/"Mandated").
  const date = /(?<![A-Za-z])(?:Dated|Dt)\b\.?\s*:?\s*(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})/i.exec(joined)
  if (date) result.noticeDate = date[1]

  return result
}

// Infers the Monitoring Format "Type of work" category from a work's free-text
// name, for sheets where that column is blank or needs a second opinion.
// Rules and their priority order were derived from a real 7-circle, 838-row
// Monitoring Format workbook (Quthbullapur Zone, Sep 2026): for each category
// below, the regex is whatever phrase actually appeared in that category's
// rows, ordered narrowest-first so a row that is genuinely "Maintenance of
// BT/CC roads" doesn't fall through to the generic BT ROADS/CC ROADS rule
// just because it happens to mention "BT" or "CC".
export const WORK_TYPE_CATEGORIES = [
  'FESTIVAL',
  'GRAVE YARDS',
  'HALLS',
  'SPORTS',
  'DESILTING WORKS',
  'JUNCTION/MEDIAN',
  'SWD',
  'UGD',
  'PARK/OPEN SPACE',
  'MAINTENANCE (BT/CC/UGD/OFFICE)',
  'BT ROADS',
  'CC ROADS'
] as const

export type WorkTypeCategory = (typeof WORK_TYPE_CATEGORIES)[number]

const RULES: [WorkTypeCategory, RegExp][] = [
  // Any "Maintenance of BT/CC roads [and storm water drains/…]" work is
  // logged under the umbrella MAINTENANCE category, never under the
  // specific thing being maintained (BT ROADS, CC ROADS, SWD, …) — so the
  // literal word "maintenance" is checked first, ahead of every other rule.
  ['MAINTENANCE (BT/CC/UGD/OFFICE)', /\bmaintenance\b/i],
  ['FESTIVAL', /festival|bonalu|bathukamma|sankranti|dasara|ganesh\s*(idol|nimajjanam|utsav|immersion)|ramzan|bakrid|urs-?e-?sandal|prayers/i],
  // "Smruthi Vanam" is a memorial garden attached to a crematorium — a named
  // exception, not inferable from "graveyard"/"crematorium" wording alone.
  // "Gravyard"/"Grayard" are the sheet's own misspellings, not alternate words.
  ['GRAVE YARDS', /grave\s*yard|graveyard|gravyard|grayard|crematorium|smashana|burning\s*plat(form)?|smruthi\s*vanam/i],
  ['HALLS', /\bhall\b|bhavan|dhobhi\s*ghat/i],
  ['SPORTS', /\bsport|stadium|skating\s*rink|cricket|badminton|municipal\s*ground/i],
  ['DESILTING WORKS', /desilt|bucket\s*cleaning/i],
  // "Widening" names are logged under the road type being widened (BT/CC),
  // UNLESS the work is really a footpath/junction-improvement job that
  // happens to mention "road widening" in passing — so both rules back off
  // when "footpath" or "junction improvement" is present, leaving those to
  // the JUNCTION/MEDIAN rule below.
  [
    'BT ROADS',
    /^(?!.*\bfootpath\b)(?!.*junction\s*improvement)(?=.*\bwidening\b)(?=.*\bbt\s*road\b)/i
  ],
  ['CC ROADS', /^(?!.*\bfootpath\b)(?!.*junction\s*improvement)(?=.*\bwidening\b)/i],
  // A row naming an explicit road type ("CC road"/"BT road") is logged under
  // that road type even when it also mentions a junction/median as a
  // waypoint or incidental landmark — so this rule backs off whenever that
  // phrasing is present, deferring to the BT/CC ROADS rules below.
  [
    'JUNCTION/MEDIAN',
    /^(?!.*\b(?:cc|bt)\s*(?:road|raod)\b)(?=.*\b(?:junction|junciton|median|lane\s*marking|road\s*safety|single\s*arm\s*pole|lighting\s*pole|footpath))/i
  ],
  // "Repair"/"repairs" is checked ahead of the specific-infrastructure rules
  // below (UGD, SWD, CC-combo, PARK) — a repair to a UGD line or a nala is
  // still logged as plain maintenance, not under what's being repaired —
  // but backs off for a park/gym/open-space, where "Repairs to park/GYM"
  // is logged under PARK/OPEN SPACE, not generic maintenance.
  ['MAINTENANCE (BT/CC/UGD/OFFICE)', /^(?!.*\b(?:park|gym|open\s*space)\b).*\brepair/i],
  [
    'SWD',
    /storm\s*water|strom\s*water|\bswg\b|\bswd\b|box\s*(drain|culvert)|\bnala\b|\bculvert\b|retaining\s*wall|rcc\s*pipe|water\s*stagnation/i
  ],
  // "Laying of UGD and CC road" names the pipe laid *under* a road that's
  // then relaid on top — in practice logged as the road type (CC ROADS),
  // not UGD, so this combined phrasing is checked before the plain UGD rule.
  ['CC ROADS', /ugd\s*(and|&)?\s*cc\s*road/i],
  ['UGD', /\bugd\b|under\s*ground\s*drain|sewer(age)?/i],
  [
    'PARK/OPEN SPACE',
    /open\s*space|\bpark\b|\bgym\b|walking\s*track|play\s*(ground|equipment)|tree\s*park|jogging\s*track/i
  ],
  [
    'MAINTENANCE (BT/CC/UGD/OFFICE)',
    /catchpit|procurement|painting|\btoilet|drinking\s*water|supply\s*(and|&)\s*(fixing|delivery)/i
  ],
  ['BT ROADS', /\bbt\s*(road|patch)|recarpet|black\s*top|strengthening/i],
  ['CC ROADS', /cc\s*(road|raod|patch(es)?|pacthes|restoration)/i]
]

// Work names describe their item type up front ("Laying of CC road from X to
// Y …") and then name landmarks along the way ("… from GHMC Park to …") that
// can accidentally match a *different* category's keyword (e.g. "Park"). So
// rules are tried first against just the leading clause — the text before
// the first "from/at/in/near/beside/under/opp." or comma — where the actual
// item type lives, and only fall back to scanning the full name (original
// behaviour) if nothing in that clause matches.
const LEAD_CLAUSE_BOUNDARY = /\b(from|at|in|near|beside|behind|opp\.?|under|for|of\s+the)\b|,/i

function leadClause(name: string): string {
  const m = LEAD_CLAUSE_BOUNDARY.exec(name)
  return m && m.index > 0 ? name.slice(0, m.index) : name
}

/**
 * Returns the best-guess "Type of work" category for a work name, or `null`
 * only when the name is blank. Rules are tried in `RULES` order and the
 * first match wins — order encodes category specificity, not alphabetical/
 * display order (see `WORK_TYPE_CATEGORIES` for the display order). Checked
 * against the leading clause before falling back to the full name — see
 * `leadClause` above. A name that matches none of the specific rules (e.g.
 * admin/office/procurement work with no distinctive keyword) defaults to
 * MAINTENANCE (BT/CC/UGD/OFFICE) — the catch-all category that, in the real
 * workbook this was validated against, accounts for most unmatched rows.
 */
export function inferTypeOfWork(workName: string): WorkTypeCategory | null {
  const name = workName.trim()
  if (!name) return null
  const lead = leadClause(name)
  for (const [category, pattern] of RULES) {
    if (pattern.test(lead)) return category
  }
  for (const [category, pattern] of RULES) {
    if (pattern.test(name)) return category
  }
  return 'MAINTENANCE (BT/CC/UGD/OFFICE)'
}

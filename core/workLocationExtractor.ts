// Best-effort extraction of a short "location" phrase from a work's
// free-text name, for the Bid Document save-name ("Bid Document: <type> @
// <location> <amount> Lakhs"). Unlike inferTypeOfWork's fixed category
// list (core/workTypeClassifier.ts), locations are unbounded free text —
// this is a heuristic, not a classifier, and can miss or pick a secondary
// landmark on unusual phrasing.

// Verb/type-noun/connector vocabulary that opens almost every work name
// ("Laying of CC Road …", "Repairs to UGD manhole & pipeline …") — this is
// stripped word-by-word off the front so whatever named landmark follows
// (often with NO preposition in front of it: "...CC Road Kennedy School to
// Bhaskar Diamond in ward...") survives, instead of only recognising
// landmarks that happen to follow "from/at/near/beside".
const LEADING_WORK_TOKENS = new Set([
  'laying', 'raod', 'construction', 'constructing', 'repair', 'repairs', 'renovation', 'renovating',
  'providing', 'provision', 'provide', 'supply', 'supplying', 'installation', 'installing',
  'development', 'developing', 'removal', 'removing', 'estimate', 'strengthening', 'desilting',
  'engaging', 'attending', 'recarpeting', 'recarpating', 'recarpet', 'balance', 'bal', 'work',
  'compost', 'pit', 'fencing', 'bucket', 'cleaning', 'machine', 'chain', 'link', 'mesh', 'open',
  'space', 'spaces', 'gym', 'gyms', 'highmast', 'light', 'lights', 'gravel', 'wmm', 'power',
  'bore', 'bores', 'pumpset', 'pumpsets', 'manhole', 'manholes', 'chamber', 'chambers',
  'pipeline', 'pipelines', 'drain', 'drains', 'drainage', 'drainges', 'katcha', 'nala', 'nalas',
  'storm', 'water', 'stagnation', 'swg', 'swd', 'ugd', 'cc', 'bt', 'road', 'roads', 'patch',
  'patches', 'pacthes', 'restoration', 'line', 'lines', 'maintenance', 'existing', 'damaged',
  'badly', 'children', 'playing', 'equipment', 'procurement', 'painting', 'catchpit', 'toilet',
  'toilets', 'drinking', 'delivery', 'fixing', 'of', 'to', 'and', '&', 'by', 'at', 'in', 'from'
])

// Same boundary workTypeClassifier's leadClause() uses to find where a
// "what's being done" clause plausibly ends — kept here as a secondary,
// looser signal alongside stripLeadingWorkTokens/WARD_CLAUSE: a name can
// lack both a recognised leading verb (arbitrary office-work phrasing,
// e.g. "Procurement of Computer, Printers...") and a ward clause (a
// zone/circle-wide work with no ward of its own) while still carrying a
// comma-delimited trail ending in a real zone/circle name worth keeping
// (e.g. "...Chintal-Circle-54, Quthbullapur Zone, CMC" -> "Quthbullapur
// Zone"). Used only to decide whether to bail to null, not to locate the
// clause itself.
const ANY_CLAUSE_SIGNAL = /\b(from|at|in|near|beside|behind|opp\.?|under|for)\b|,/i

// Cuts off the "ward no 123" / "ward no: 123-Suraram" tail — whatever
// comes after that is circle/zone/corporation boilerplate, not location.
const WARD_CLAUSE = /\bward\s*(?:no\.?|number)?\s*[:.\-]?\s*\d+/i

// House/plot/survey numbering glued directly onto a place name (e.g.
// "plot no.154 Sri Venkateswara residency") — stripped as a prefix rather
// than rejecting the whole segment, so the actual name after it survives.
// "road no" is deliberately NOT included here — a road number ("Road
// No.10", "road no:10 And 10C") is itself a valid location, not noise
// (see the "road"-lookahead exception in stripLeadingWorkTokens below).
const IDENTIFIER_PREFIX = /^(?:plot\s*no\.?|p\.?no\.?|h\.?no\.?|hno|s\.?no\.?|sy\.?\s*no\.?|syno|survey\s*no\.?)\s*[:.\-]?\s*[\d/.-]*\s*/i

// Whole-segment noise: corporation/zone/scheme abbreviations that
// occasionally end up as the last split segment when nothing more
// specific was named (e.g. "...in Jeedimetla Circle-55, Quthbullapur
// Zone, CMC" with no colony name) — not places, so dropped rather than
// surfaced as a misleading "location".
const NON_PLACE_WHOLE = /^(cmc|ghmc|nmc|mmc|mgf|circle[-\s]?\d*|under)\b/i

const NUMERIC_ONLY = /^[\d\s\-/.,]*$/

const SPLIT = /,|\bfrom\b|\bto\b|\bvia\b|\bbeside\b|\bnear\b|\bopp\.?\b|\bat\b|\bin\b|\bof\b/i

// Scheme/fund and reservation-category boilerplate that tags along after a
// real place name ("Nizampet Municipal Corporation under Municipal General
// Funds 2025-26", "...2025-26 (Reserved for SC"). Stripped off the tail of
// a segment rather than rejecting it outright, since the place name in
// front of it is still a usable location.
const BOILERPLATE_SUFFIX =
  /\s*\(?(?:reserved\s+for\s+[a-z\s/]+?\)?|under\s+municipal\s+general\s+funds?\s*\d{0,4}(?:-\d{2,4})?)\s*$/i

function cleanSegment(raw: string): string {
  let s = raw.trim().replace(/^[.,:\-()]+|[.,:\-()]+$/g, '').trim()
  let m = IDENTIFIER_PREFIX.exec(s)
  while (m) {
    s = s.slice(m[0].length).replace(/^[.,:\-()]+|[.,:\-()]+$/g, '').trim()
    m = IDENTIFIER_PREFIX.exec(s)
  }
  let suf = BOILERPLATE_SUFFIX.exec(s)
  while (suf) {
    s = s.slice(0, suf.index).replace(/^[.,:\-()]+|[.,:\-()]+$/g, '').trim()
    suf = BOILERPLATE_SUFFIX.exec(s)
  }
  return s
}

/**
 * Consumes `LEADING_WORK_TOKENS` word-by-word off the front of `name`,
 * stopping at the first word that isn't recognised (almost always the
 * start of an actual landmark/colony name). "road" is a special case:
 * normally type-noun filler ("CC Road"), but when the NEXT word is "no"
 * ("road no 5", "road no:10") the pair is itself the location, so the
 * loop stops right before "road" instead of consuming it.
 */
function stripLeadingWorkTokens(name: string): string {
  let i = 0
  const n = name.length
  while (i < n) {
    while (i < n && /[\s,]/.test(name[i])) i++
    const rest = name.slice(i)
    const m = /^[A-Za-z.&]+/.exec(rest)
    if (!m) break
    const word = m[0].toLowerCase().replace(/\.+$/, '')
    if (word === 'road') {
      const after = rest.slice(m[0].length)
      const nextM = /^[\s,]*([A-Za-z.&]+)/.exec(after)
      const nextWord = nextM ? nextM[1].toLowerCase().replace(/\.+$/, '') : ''
      if (nextWord === 'no') break
    }
    if (!LEADING_WORK_TOKENS.has(word)) break
    i += m[0].length
  }
  return name.slice(i)
}

/**
 * Best-guess location phrase for a work name (e.g. "Venkateshwaranagar",
 * "road no:10 And 10C", "MN Reddy Nagar colony"), or `null` when nothing
 * plausible was found. Strips the leading verb/type-noun clause (see
 * stripLeadingWorkTokens), then takes the LAST non-noise segment before
 * the "ward no" clause — names typically run "<what's being done> <one or
 * more landmarks, often chained with from/to> in ward no <N>", and the
 * landmark closest to the ward clause is the actual work site, not an
 * incidental waypoint named earlier.
 *
 * Returns null when the leading-token strip consumed nothing, no ward
 * clause was found, and the name carries no other clause-like structure
 * (comma/preposition) either — i.e. nothing in the name anchors where a
 * location clause might start (e.g. a bare "Annual desilting").
 */
export function inferWorkLocation(workName: string): string | null {
  const name = workName.trim()
  if (!name) return null

  const stripped = stripLeadingWorkTokens(name)
  const consumed = stripped.length < name.length
  const ward = WARD_CLAUSE.exec(stripped)
  if (!consumed && !ward && !ANY_CLAUSE_SIGNAL.test(name)) return null

  const extractFrom = (part: string): string[] =>
    part
      .split(SPLIT)
      .map(cleanSegment)
      .filter((s) => s && !NUMERIC_ONLY.test(s) && !NON_PLACE_WHOLE.test(s))

  const locPart = ward ? stripped.slice(0, ward.index) : stripped
  const segments = extractFrom(locPart)
  if (segments.length > 0) return segments[segments.length - 1]

  // "ward-294 Dundigal Circle-59,Quthbullapur Zone,CMC" — the ward number is
  // glued directly onto "ward" with no landmark before it at all, so the
  // pre-ward half is empty. Fall back to whatever follows the ward number
  // instead of giving up. Here the FIRST segment wins, not the last: the
  // name right after the ward number is the specific colony/village tied to
  // that ward ("Dundigal", "Bachupally"), while later segments trail off
  // into the broader circle/zone/corporation it sits in — real but far less
  // specific, so only used when nothing more local is named at all.
  if (ward) {
    const after = extractFrom(stripped.slice(ward.index + ward[0].length))
    if (after.length > 0) return after[0]
  }

  return null
}

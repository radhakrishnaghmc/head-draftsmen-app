/**
 * The CMC (Cyberabad Municipal Corporation) zone ↔ circle directory,
 * hard-coded from the same private credentials sheet the login checks against
 * (the "passwords link"): every circle, its parent zone, and its circle
 * number (CNO). It's a small, fixed list that changes only when the
 * corporation reorganises, so baking it into the app (rather than fetching
 * it) keeps zone inference working offline and for every user regardless of
 * which circle they logged in as.
 *
 * Why it exists: a work's "Name of the work" almost always mentions its
 * *circle* (e.g. "…road at Miyapur…") but never spells out its *zone*
 * ("Serilingampally"). So the login-name match in enforceZoneCircle can fill
 * the Circle column from the name but leaves Zone blank. This directory
 * bridges that gap — given the circle, we know the zone (and the CNO).
 */

export interface ZoneCircleEntry {
  zone: string
  circle: string
  /** Circle number (CNO). */
  cno: string
}

export const CMC_ZONE_CIRCLES: ZoneCircleEntry[] = [
  // Kukatpally zone
  { zone: 'Kukatpally', circle: 'Madhapur', cno: '50' },
  { zone: 'Kukatpally', circle: 'Allwyn Colony', cno: '51' },
  { zone: 'Kukatpally', circle: 'Kukatpally', cno: '52' },
  { zone: 'Kukatpally', circle: 'Moosapet', cno: '53' },
  // Quthbullapur zone
  { zone: 'Quthbullapur', circle: 'Chintal', cno: '54' },
  { zone: 'Quthbullapur', circle: 'Jeedimetla', cno: '55' },
  { zone: 'Quthbullapur', circle: 'Kompally', cno: '56' },
  { zone: 'Quthbullapur', circle: 'Gajularamaram', cno: '57' },
  { zone: 'Quthbullapur', circle: 'Nizampet', cno: '58' },
  { zone: 'Quthbullapur', circle: 'Dundigal', cno: '59' },
  { zone: 'Quthbullapur', circle: 'Medchal', cno: '60' },
  // Serilingampally zone
  { zone: 'Serilingampally', circle: 'Narsingi', cno: '45' },
  { zone: 'Serilingampally', circle: 'Patancheruvu', cno: '46' },
  { zone: 'Serilingampally', circle: 'Ameenpur', cno: '47' },
  { zone: 'Serilingampally', circle: 'Miyapur', cno: '48' },
  { zone: 'Serilingampally', circle: 'Serilingampally', cno: '49' }
]

const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/**
 * What a work name (or an existing Circle cell) resolved to against the
 * directory. Values are only set when they're *unambiguous*: a name that
 * mentions two different circles from two different zones (a road running
 * between circles) leaves `zone` blank rather than guessing, and one that
 * mentions two circles at all leaves `circle` blank — but if those circles
 * share a zone, that zone is still returned.
 */
export interface DirectoryMatch {
  zone?: string
  circle?: string
  cno?: string
}

/**
 * Resolve the zone/circle a work name belongs to against the directory.
 *
 * A work name usually carries an explicit "<name> circle, <name> zone"
 * designation (e.g. "…in Nizampet circle, Quthbullapur zone") *alongside* the
 * road's own endpoints/locality, which are often other places' names — e.g.
 * "laying of Miyapur to Bachupally road … in Nizampet circle" mentions Miyapur
 * (a Serilingampally circle) purely as a road endpoint. So the "<name> circle"
 * / "<name> zone" phrasing is matched first and wins outright; only when a name
 * carries neither do we fall back to scanning for a bare circle name anywhere
 * in the text (and then only commit to an unambiguous single circle/zone).
 */
export function resolveFromDirectory(text: string | undefined): DirectoryMatch {
  const hay = (text ?? '').trim()
  if (!hay) return {}

  // 1. Explicit "<name> circle" / "<name> zone" designation wins over any
  //    incidental place names elsewhere in the string.
  const taggedCircle = CMC_ZONE_CIRCLES.find((e) =>
    new RegExp(`\\b${escapeRegExp(e.circle)}\\s+circle\\b`, 'i').test(hay)
  )
  const distinctZoneNames = [...new Set(CMC_ZONE_CIRCLES.map((e) => e.zone))]
  const taggedZone = distinctZoneNames.find((z) => new RegExp(`\\b${escapeRegExp(z)}\\s+zone\\b`, 'i').test(hay))

  if (taggedCircle || taggedZone) {
    // A tagged circle settles the zone via the directory (authoritative), so a
    // separately tagged zone is only needed when no circle was tagged.
    return {
      zone: taggedCircle ? taggedCircle.zone : taggedZone,
      circle: taggedCircle?.circle,
      cno: taggedCircle?.cno
    }
  }

  // 2. No explicit designation — fall back to a bare scan of any known circle
  //    name in the text, committing only to an unambiguous single circle/zone.
  const matched = CMC_ZONE_CIRCLES.filter((e) =>
    new RegExp(`\\b${escapeRegExp(e.circle)}\\b`, 'i').test(hay)
  )
  if (matched.length === 0) return {}

  const distinctCircles = [...new Set(matched.map((e) => e.circle.toLowerCase()))]
  const distinctZones = [...new Set(matched.map((e) => e.zone.toLowerCase()))]

  const single = distinctCircles.length === 1 ? matched.find((e) => e.circle.toLowerCase() === distinctCircles[0]) : undefined
  const zone = distinctZones.length === 1 ? matched.find((e) => e.zone.toLowerCase() === distinctZones[0])!.zone : undefined

  return { zone, circle: single?.circle, cno: single?.cno }
}

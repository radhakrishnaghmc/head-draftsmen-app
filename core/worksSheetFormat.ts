import type { PlaceholderMatch } from './createDocument'

/**
 * Reading an office's own Works List sheet. Every office keeps its works in its own layout with hand-typed headers
 * ("Cirlce no. and Name", "Techincal Sanction No.", "E C V"), and a fuzzy semantic match alone mis-files some of them
 * (e.g. a Technical Sanction date read as the Admin Sanction date). So headers are first recognised by name here
 * (spelling slips and spacing tolerated); whatever this doesn't recognise is left for the fuzzy matcher. The importer
 * keeps the sheet's own columns, names and order — only the app's internal name for each column is decided.
 */

const norm = (h: string) =>
  h
    .replace(/[\r\n]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()

/** Sheet header (normalized) -> the Works List column the app reads it as. First match per Works List column wins. */
const HEADER_ALIASES: [RegExp, string][] = [
  [/^win ?code$/, 'Wincode'],
  [/^(c[a-z]{3,5}\s*(no\.?|number)|cno)$/, 'Circle number'],
  [/^c(ir|ri)[a-z]{1,3}\b/, 'Circle'],
  [/^name of (the )?work/, 'Name of the work'],
  [/^(amount of )?estimate (amount|cost|value)|^amount of estimate$/, 'Amount of estimate'],
  [/^e ?c ?v\b|^estimate(d)? contract value/, 'ECV'],
  [/^(nit|tender notice) ?(no|number)|^tender notice no/, 'Tender Notice No'],
  [/^contract (value|amount)/, 'Contract Amount'],
  [/tender percentage|^% ?(above|below|less)|^tender (premium|excess)/, 'Tender Percentage'],
  [/^name of (the )?(contractor|agency)|^agency name|^contractor( name)?$/, 'Name of the Agency'],
  [/phone|mobile|contact (no|number)/, 'Phone number of the agency'],
  [/^te[a-z]{1,6}cal sanction (no|number)|^ts no/, 'Technical Sanc No'],
  [/^te[a-z]{1,6}cal sanction (date|dt)|^ts (date|dt)/, 'TS date'],
  [/^admin(istrative)? sanction (date|dt)|^as (date|dt)/, 'Admin Sanction Date'],
  [/^as given by|^sanction(ed)? by|^admin(istrative)? sanction (by|authority)/, 'Sanction By'],
  [/^poc\b|period of completion|^completion period|^time of completion/, 'Completion Period'],
  [/^loa date|^intimation date|^date of loa/, 'Intimation Date'],
  [/^agreement (dt|date)|^agmt (dt|date)/, 'Agmt Date'],
  [/^(fy|financial year)$/, 'Financial Year']
]

/**
 * Maps each Works List column to the sheet header that feeds it, recognised by name. A column the sheet doesn't carry
 * (or names in a way not recognised here) comes back with `column: null` so the caller can try the fuzzy matcher on
 * what's left (see `unclaimedHeaders`).
 */
export function aliasMapping(labels: string[], headers: string[]): PlaceholderMatch[] {
  const claimed = new Map<string, string>() // works column -> sheet header
  const used = new Set<string>()
  for (const h of headers) {
    const n = norm(h)
    const hit = HEADER_ALIASES.find(([re, col]) => !claimed.has(col) && re.test(n))
    if (hit && !used.has(h)) {
      claimed.set(hit[1], h)
      used.add(h)
    }
  }
  return labels.map((label) => ({ label, column: claimed.get(label) ?? null, score: claimed.has(label) ? 1 : 0 }))
}

/** Combine the name-based mapping with a fuzzy one made over only the headers the name-based pass left unclaimed. */
export function mergeMappings(byName: PlaceholderMatch[], fuzzy: PlaceholderMatch[]): PlaceholderMatch[] {
  const fuzzyFor = new Map(fuzzy.map((m) => [m.label, m]))
  return byName.map((m) => (m.column ? m : (fuzzyFor.get(m.label) ?? m)))
}

/** The headers `aliasMapping` did not claim — what the fuzzy matcher is allowed to work with. */
export function unclaimedHeaders(byName: PlaceholderMatch[], headers: string[]): string[] {
  const taken = new Set(byName.map((m) => m.column).filter((c): c is string => !!c))
  return headers.filter((h) => !taken.has(h))
}

const isBanner = (row: Record<string, string>) => {
  const filled = Object.values(row).filter((v) => (v ?? '').trim() !== '')
  return filled.length === 1 && /^(\d{4}\s*-\s*\d{2,4}|(fy|f\.y\.?)\s*\d{4}.*)$/i.test(filled[0].trim())
}

/** "Rs.61,67,723.45" -> "6167723.45"; plain numbers and blanks are left as they are. */
function plainRupees(v: string): string {
  const s = v.trim()
  if (!/[a-z]/i.test(s)) return v
  const n = s.replace(/^(rs|inr)\.?\s*/i, '').replace(/,/g, '')
  return /^\d+(\.\d+)?$/.test(n) ? n : v
}

/**
 * Cleans an office sheet's rows before the schema is applied: section banner rows ("2025-26") and rows with neither a
 * work name nor a Wincode are not works, and a contract value typed as "Rs.61,67,723.45" becomes a plain number.
 * Everything else stays exactly as typed so the database reads like the sheet.
 */
export function cleanSheetRows(headers: string[], rows: Record<string, string>[]): Record<string, string>[] {
  const nameH = headers.find((h) => /^name of (the )?work/.test(norm(h)))
  const codeH = headers.find((h) => /^win ?code$/.test(norm(h)))
  const valueH = headers.find((h) => /^contract (value|amount)/.test(norm(h)))
  // Without a recognisable name/Wincode column there is nothing to tell a work from a stray row: keep every row.
  const identifiable = !!(nameH || codeH)
  return rows
    .filter(
      (r) => !isBanner(r) && (!identifiable || (nameH && (r[nameH] ?? '').trim()) || (codeH && (r[codeH] ?? '').trim()))
    )
    .map((r) => (valueH && r[valueH] ? { ...r, [valueH]: plainRupees(r[valueH]) } : r))
}

/**
 * A zone-level (SE) sheet names the circle in one combined cell ("55-Jeedimetla"), and often leaves it blank, naming it
 * only inside the work ("... Medchal Circle 60 ..."). Returns the circle as the sheet writes it plus its number;
 * `circleNames` maps a circle number to its name. A filled cell is never altered.
 */
export function circleOfWork(
  circleCell: string,
  workName: string,
  circleNames: Record<string, string>
): { circle: string; cno: string } {
  const cell = circleCell.trim()
  if (cell) return { circle: cell, cno: /^(\d{1,3})\b/.exec(cell)?.[1] ?? '' }
  const cno =
    /circle[\s.:-]*(?:no\.?\s*)?(\d{1,3})\b/i.exec(workName)?.[1] ?? /\b(\d{2,3})\s*circle\b/i.exec(workName)?.[1]
  if (!cno) return { circle: '', cno: '' }
  const name = circleNames[cno]
  return { circle: name ? `${cno}-${name}` : '', cno }
}

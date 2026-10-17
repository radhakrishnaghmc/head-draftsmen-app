/**
 * A key under which spelling variants of one contractor's name collapse:
 * "M/S AHR Constructions", "M/s. A.H.R. CONSTRUCTIONS" and "AHR CONSTRUCTIONS"
 * all give the same key. Case, a leading M/S / Messrs, punctuation and spacing
 * are ignored; the words themselves are not, so genuinely different agencies
 * stay apart.
 */
export function agencyKey(name: string | undefined): string {
  return (name ?? '')
    .toLowerCase()
    .replace(/^\s*(m\s*\/\s*s\.?|messrs\.?)(?=\s|$)/, ' ')
    .replace(/[^a-z0-9]+/g, '')
}

/**
 * Merges agency spelling variants: maps each name to the variant used most
 * often (ties go to the first seen), so the same contractor always displays
 * under one name.
 */
export function canonicalAgencyNames(names: Iterable<string | undefined>): Map<string, string> {
  const counts = new Map<string, Map<string, number>>()
  for (const raw of names) {
    const name = (raw ?? '').trim()
    const key = agencyKey(name)
    if (!key) continue
    const variants = counts.get(key) ?? new Map<string, number>()
    variants.set(name, (variants.get(name) ?? 0) + 1)
    counts.set(key, variants)
  }
  const out = new Map<string, string>()
  for (const [key, variants] of counts) {
    let best = ''
    let bestCount = 0
    for (const [name, n] of variants) if (n > bestCount) [best, bestCount] = [name, n]
    out.set(key, best)
  }
  return out
}

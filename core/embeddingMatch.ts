/** Cosine similarity between two equal-length embedding vectors. */
export function cosineSimilarity(a: number[], b: number[]): number {
  let dot = 0
  let normA = 0
  let normB = 0
  for (let i = 0; i < a.length; i++) {
    dot += a[i] * b[i]
    normA += a[i] * a[i]
    normB += b[i] * b[i]
  }
  if (normA === 0 || normB === 0) return 0
  return dot / (Math.sqrt(normA) * Math.sqrt(normB))
}

export interface EmbeddingMatch {
  index: number
  score: number
}

/** Rank every candidate vector against a query vector by cosine similarity, best first. */
export function rankByEmbedding(queryVector: number[], candidateVectors: number[][]): EmbeddingMatch[] {
  return candidateVectors
    .map((v, index) => ({ index, score: cosineSimilarity(queryVector, v) }))
    .sort((a, b) => b.score - a.score)
}

/**
 * Minimum cosine similarity to treat two "Name of Work" strings as the same
 * work (Works List row matching, Schedule A, the cross-file work-name
 * consistency check). Deliberately higher than the 0.5 used for short-string
 * column/placeholder matching (core/columnMatch.ts) — full work-name
 * sentences carry a lot of shared civic boilerplate ("Laying of CC road from
 * … ward no …, … Circle-…, … Zone, …") that inflates cosine similarity between
 * two DIFFERENT works far more than it does between two short header strings.
 *
 * Measured on the bundled all-MiniLM-L6-v2 model against a real false match
 * from this app's own use (two different roads, same circle/zone/corporation
 * template): the wrong row scored 0.81 — comfortably above the old 0.5 — while
 * a genuine same-work wording-drift pair scored 0.88. 0.85 sits between them.
 *
 * This is not a complete fix: a same-template, different-street pair can
 * still score higher than a genuine paraphrase (measured as high as 0.94) —
 * cosine similarity on full sentences can't fully separate "same work,
 * reworded" from "different work, same civic phrasing". Every caller of this
 * threshold treats a match as "supporting details only, never the work's own
 * identifying fields" for exactly this reason (see deriveFields /
 * noteSubmittedFromRow's upload-first precedence) — the threshold reduces how
 * often a wrong match happens, it doesn't make a wrong match harmless on its
 * own.
 */
export const WORK_IDENTITY_MATCH_THRESHOLD = 0.85

/** The ward / circle / zone / corporation keywords of a name, in the order they appear ("ward>circle>zone>corp"). */
const tailOrder = (s: string) =>
  [...s.matchAll(/\bward\b|\bcir[ck]le\b|\bzone\b|\b(?:cmc|ghmc|mmc)\b/gi)].map((m) => (/^(cmc|ghmc|mmc)$/i.test(m[0]) ? 'corp' : m[0].toLowerCase().replace('cirkle', 'circle'))).join('>')

const trigrams = (s: string) => {
  const t = ` ${s.toLowerCase().replace(/[^a-z0-9/]+/g, ' ').trim()} `
  return new Set(Array.from({ length: Math.max(t.length - 2, 0) }, (_, i) => t.slice(i, i + 3)))
}

/**
 * Two names that open with the same work type ("Laying of CC Road") and end with the same ward / zone / corporation
 * wording in the same order are boilerplate-identical, so the embedding score says little. Only then is the shared
 * wording stripped and the place-specific parts compared (trigram Dice, typo-tolerant); a low score means different works.
 */
function sharedTemplateDiffers(a: string, b: string): boolean {
  const lead = (s: string) => s.toLowerCase().replace(/[^a-z ]+/g, ' ').split(/\s+/).filter(Boolean).slice(0, 4).join(' ')
  const tail = tailOrder(a)
  if (lead(a).split(' ').length < 4 || lead(a) !== lead(b) || !tail.includes('ward') || !tail.includes('zone') || tail !== tailOrder(b)) return false
  const x = trigrams(stripWorkBoilerplate(a))
  const y = trigrams(stripWorkBoilerplate(b))
  const common = [...x].filter((g) => y.has(g)).length
  return (2 * common) / (x.size + y.size || 1) < 0.5
}

/**
 * Whether two work names carry different ward or circle numbers ("ward no
 * 259" vs "ward no -274", "Circle-53" vs "Circle-58"). Same-template works
 * ("Making Arrangements for Bathukamma Festival at ... ward ... Circle ...")
 * score above WORK_IDENTITY_MATCH_THRESHOLD on cosine alone, so an embedding
 * match must also clear this check — a differing number on either side
 * means a different work, however similar the sentence reads.
 */
export function workNamesConflict(a: string, b: string): boolean {
  const nums = (s: string, re: RegExp): Set<string> =>
    new Set([...s.matchAll(re)].map((m) => String(Number(m[1]))))
  const differs = (re: RegExp): boolean => {
    const x = nums(a, re)
    const y = nums(b, re)
    if (x.size === 0 || y.size === 0) return false
    return ![...x].some((n) => y.has(n))
  }
  // House numbers ("H.No 6-41/1"): same-street-template works in one ward differ only by these.
  const houses = (s: string) =>
    new Set([...s.matchAll(/\bh\.?\s*no\.?\s*[-.:]?\s*(\d[\d/-]*\d|\d)/gi)].map((m) => m[1].replace(/\/$/, '')))
  const ha = houses(a)
  const hb = houses(b)
  const houseConflict = ha.size > 0 && hb.size > 0 && ![...ha].some((n) => hb.has(n))
  return (
    houseConflict ||
    sharedTemplateDiffers(a, b) ||
    differs(/ward\s*(?:no\.?|number)?\s*[-.:]?\s*(\d+)/gi) ||
    differs(/circle\s*(?:no\.?)?\s*[-.:]?\s*(\d+)/gi)
  )
}

/**
 * A work name with the wording every work shares removed (the work type — "Laying of CC Road", "Construction of", "Providing" —
 * and the ward / circle / zone / corporation tail), leaving the place-specific part (streets, house numbers, colony). Embed
 * THIS, not the full title: shared boilerplate otherwise dominates the cosine score and two different roads look identical.
 */
export function stripWorkBoilerplate(name: string): string {
  const stripped = name
    .replace(/\((?:[^)]*reserved[^)]*|\d+(?:st|nd|rd|th)\s*call)\)/gi, ' ')
    .replace(/\b(?:laying|construction|providing|development|improvements?|repairs?|renovation|formation|strengthening|widening|re-?laying|maintenance|making arrangements)\b(?:\s+(?:of|to|for|and|&|the|at))*/gi, ' ')
    .replace(/\b(?:cc|bt|wbm|rcc|pcc|c\.c\.?)\s*(?:roads?|pavements?|drains?)\b/gi, ' ')
    .replace(/\b(?:storm\s*water|sw|sewerage|sewer|water\s*supply|road|roads|drain|drains|nala|culvert|footpath|footpaths|pipe\s*line)\b/gi, ' ')
    .replace(/\bward\s*(?:no\.?|number)?\s*[-.:]?\s*\d+/gi, ' ')
    .replace(/\bcir[ck]le\s*(?:no\.?)?\s*[-.:]?\s*\d+/gi, ' ')
    .replace(/\b(?:[a-z]+\s+)?zone\b|\b(?:cmc|ghmc|mmc|kpz|ee|se)\b|\bcircle\b/gi, ' ')
    .replace(/\b(?:from|to|at|in|of|and|the|for|by|near|via)\b/gi, ' ')
    .replace(/[,.;:]+(?=\s|$)/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  return stripped || name // a title that is only boilerplate keeps its full text
}

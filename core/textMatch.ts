import type { RateEntry } from './rateDatabase'

function tokenize(s: string): string[] {
  return s
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter((t) => t.length > 0)
}

function termFrequency(tokens: string[]): Map<string, number> {
  const tf = new Map<string, number>()
  for (const t of tokens) tf.set(t, (tf.get(t) ?? 0) + 1)
  return tf
}

export interface RateIndex {
  entries: RateEntry[]
  idf: Map<string, number>
  docTf: Map<string, number>[]
  docNorm: number[]
}

/**
 * Build a TF-IDF cosine-similarity index over a rate database's descriptions.
 * IDF weighting matters here because these descriptions share huge amounts
 * of boilerplate phrasing ("including cost and conveyance of all materials…
 * complete for finished item of work") across otherwise-unrelated items —
 * plain word-overlap would be dominated by that boilerplate. Terms that
 * appear in most documents get a low weight automatically; the distinctive
 * terms (materials, dimensions, standard clause numbers) drive the score.
 */
export function buildRateIndexForMatching(entries: RateEntry[]): RateIndex {
  const docTokens = entries.map((e) => tokenize(e.description))
  const df = new Map<string, number>()
  for (const tokens of docTokens) {
    for (const t of new Set(tokens)) df.set(t, (df.get(t) ?? 0) + 1)
  }
  const n = entries.length
  const idf = new Map<string, number>()
  for (const [t, count] of df) idf.set(t, Math.log((n + 1) / (count + 1)) + 1)

  const docTf = docTokens.map((tokens) => termFrequency(tokens))
  const docNorm = docTf.map((tf) => {
    let sumSq = 0
    for (const [t, f] of tf) {
      const w = f * (idf.get(t) ?? 0)
      sumSq += w * w
    }
    return Math.sqrt(sumSq)
  })

  return { entries, idf, docTf, docNorm }
}

export interface MatchResult {
  entry: RateEntry
  score: number
}

/** Find every rate-database entry for a query description, best score first. */
export function findMatches(query: string, index: RateIndex): MatchResult[] {
  const queryTf = termFrequency(tokenize(query))
  let queryNormSq = 0
  const queryWeights = new Map<string, number>()
  for (const [t, f] of queryTf) {
    const w = f * (index.idf.get(t) ?? 0)
    queryWeights.set(t, w)
    queryNormSq += w * w
  }
  const queryNorm = Math.sqrt(queryNormSq)
  if (queryNorm === 0) return []

  const results: MatchResult[] = []
  for (let i = 0; i < index.entries.length; i++) {
    const docNorm = index.docNorm[i]
    if (docNorm === 0) continue
    let dot = 0
    for (const [t, qw] of queryWeights) {
      const docTf = index.docTf[i].get(t)
      if (!docTf) continue
      dot += qw * (docTf * (index.idf.get(t) ?? 0))
    }
    if (dot === 0) continue
    results.push({ entry: index.entries[i], score: dot / (queryNorm * docNorm) })
  }
  return results.sort((a, b) => b.score - a.score)
}

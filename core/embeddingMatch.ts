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

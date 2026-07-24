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

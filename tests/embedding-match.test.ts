import { describe, it, expect } from 'vitest'
import { cosineSimilarity, rankByEmbedding } from '../core/embeddingMatch'

describe('cosineSimilarity', () => {
  it('is 1 for identical vectors', () => {
    expect(cosineSimilarity([1, 2, 3], [1, 2, 3])).toBeCloseTo(1, 10)
  })

  it('is 0 for orthogonal vectors', () => {
    expect(cosineSimilarity([1, 0], [0, 1])).toBeCloseTo(0, 10)
  })

  it('is -1 for opposite vectors', () => {
    expect(cosineSimilarity([1, 2], [-1, -2])).toBeCloseTo(-1, 10)
  })

  it('is invariant to vector magnitude (only direction matters)', () => {
    expect(cosineSimilarity([1, 2, 3], [2, 4, 6])).toBeCloseTo(1, 10)
  })

  it('returns 0 for a zero vector rather than dividing by zero', () => {
    expect(cosineSimilarity([0, 0, 0], [1, 2, 3])).toBe(0)
  })
})

describe('rankByEmbedding', () => {
  it('ranks candidates by similarity to the query, best first', () => {
    const query = [1, 0, 0]
    const candidates = [
      [0, 1, 0], // orthogonal, score 0
      [1, 0.1, 0], // close to query, high score
      [-1, 0, 0] // opposite, score -1
    ]
    const ranked = rankByEmbedding(query, candidates)
    expect(ranked.map((r) => r.index)).toEqual([1, 0, 2])
    expect(ranked[0].score).toBeGreaterThan(ranked[1].score)
    expect(ranked[1].score).toBeGreaterThan(ranked[2].score)
  })

  it('returns an empty list for an empty candidate set', () => {
    expect(rankByEmbedding([1, 0], [])).toEqual([])
  })
})

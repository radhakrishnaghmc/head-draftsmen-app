import { describe, it, expect } from 'vitest'
import { cosineSimilarity, rankByEmbedding, WORK_IDENTITY_MATCH_THRESHOLD } from '../core/embeddingMatch'

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

describe('WORK_IDENTITY_MATCH_THRESHOLD', () => {
  // Pins the threshold against the actual cosine scores measured on the
  // bundled all-MiniLM-L6-v2 model for a real bug: "Laying of CC Road from
  // RGK STP to Children Park … Nizampet Circle-58 …" was wrongly matched to
  // the Works List's unrelated "Laying of CC road from Shubamkaree temple to
  // babu Jagjivan park … Nizampet circle-58 …" (score 0.81) under the old 0.5
  // threshold, while a genuine same-work wording-drift pair scored 0.88. If
  // this threshold ever drifts back down, this is the regression it exists
  // to prevent.
  it('rejects a same-boilerplate-different-work score (~0.81, the real false match)', () => {
    const a = [1, 0]
    const b = [0.81, 0.586] // cosine(a, b) ≈ 0.81
    expect(cosineSimilarity(a, b)).toBeLessThan(WORK_IDENTITY_MATCH_THRESHOLD)
  })

  it('accepts a genuine wording-drift score (~0.88, the real true match)', () => {
    const a = [1, 0]
    const b = [0.88, 0.475] // cosine(a, b) ≈ 0.88
    expect(cosineSimilarity(a, b)).toBeGreaterThan(WORK_IDENTITY_MATCH_THRESHOLD)
  })

  it('is set above the old 0.5 default that let the real false match through', () => {
    expect(WORK_IDENTITY_MATCH_THRESHOLD).toBeGreaterThan(0.5)
  })
})

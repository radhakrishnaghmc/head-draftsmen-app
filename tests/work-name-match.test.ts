import { describe, expect, it } from 'vitest'
import { compareWorkNames, normWorkName, workNameMismatchMessage } from '../core/workNameMatch'

describe('normWorkName', () => {
  it('drops punctuation, collapses spaces and lowercases', () => {
    expect(normWorkName('  Road  from  A—to—B. ')).toBe('road from a to b')
  })
  it('returns empty for undefined/blank', () => {
    expect(normWorkName(undefined)).toBe('')
    expect(normWorkName('   ')).toBe('')
  })
})

describe('compareWorkNames (no embeddings — word-overlap fallback)', () => {
  it('is a match when the names are identical apart from punctuation/case', () => {
    expect(compareWorkNames('Road from A to B', 'road from a to b.').status).toBe('match')
  })

  it('is a match on minor wording differences for the same work', () => {
    const r = compareWorkNames(
      'Construction of CC road from village A to village B',
      'Construction of CC road from A to B'
    )
    expect(r.status).toBe('match')
  })

  it('is a mismatch for two clearly different works', () => {
    const r = compareWorkNames(
      'Construction of CC road from A to B',
      'Improvements to drainage system near market yard'
    )
    expect(r.status).toBe('mismatch')
  })

  it("is 'unknown' when either name is blank — nothing to compare, don't block", () => {
    expect(compareWorkNames('', 'Road from A to B').status).toBe('unknown')
    expect(compareWorkNames('Road from A to B', undefined).status).toBe('unknown')
  })
})

describe('compareWorkNames (with embeddings)', () => {
  it('uses cosine similarity when the exact match fails: close vectors -> match', () => {
    const r = compareWorkNames('Work A worded one way', 'Work A worded another way', {
      aVector: [1, 0, 0],
      bVector: [0.9, 0.1, 0]
    })
    expect(r.status).toBe('match')
  })

  it('distant vectors -> mismatch', () => {
    const r = compareWorkNames('Work A', 'Work B', {
      aVector: [1, 0, 0],
      bVector: [0, 1, 0]
    })
    expect(r.status).toBe('mismatch')
  })
})

describe('workNameMismatchMessage', () => {
  it('names both works and asks for the same work', () => {
    const msg = workNameMismatchMessage('Work B', 'Work A')
    expect(msg).toContain('Work B')
    expect(msg).toContain('Work A')
    expect(msg.toLowerCase()).toContain('same work')
  })
})

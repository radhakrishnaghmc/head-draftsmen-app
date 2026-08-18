import { cosineSimilarity, WORK_IDENTITY_MATCH_THRESHOLD } from './embeddingMatch'

/**
 * Decide whether two "Name of Work" strings describe the same work — used to
 * block a workspace from mixing files for different works (e.g. an
 * estimate/BOQ uploaded for a different work than the tender the Work Order /
 * Agreement / Schedule A is being built for).
 *
 * Work names for the same work rarely match character-for-character across
 * files — a BOQ's title block, the L-1 selection form and the Works List each
 * word it slightly differently (abbreviations, punctuation, rephrasing). So an
 * exact normalized match is only the fast path: when it fails, the caller can
 * supply embedding vectors for the two names and the cosine similarity decides
 * against WORK_IDENTITY_MATCH_THRESHOLD (core/embeddingMatch.ts) — higher than
 * the threshold used for short-string column/placeholder matching, because
 * full work-name sentences share enough civic boilerplate that two DIFFERENT
 * works can otherwise score as a match. Without embeddings we fall back to
 * word-overlap (Jaccard) so trivial wording differences still count as the
 * same work while a genuinely different work is flagged. When either name is
 * blank there's nothing to compare and the status is 'unknown' (the caller
 * shouldn't block).
 */

// Word-overlap fallback when no embeddings are available.
const OVERLAP_THRESHOLD = 0.5

export type WorkNameStatus = 'match' | 'mismatch' | 'unknown'

export interface WorkNameCompare {
  status: WorkNameStatus
  /** Similarity that settled it (1 for an exact match), for debugging/hints. */
  score?: number
}

/** Normalize a work name for comparison: drop punctuation, collapse spaces, lowercase. */
export function normWorkName(s: string | undefined): string {
  return (s ?? '')
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
}

export function compareWorkNames(
  a: string | undefined,
  b: string | undefined,
  embeddings?: { aVector: number[]; bVector: number[] }
): WorkNameCompare {
  const na = normWorkName(a)
  const nb = normWorkName(b)
  if (!na || !nb) return { status: 'unknown' }
  if (na === nb) return { status: 'match', score: 1 }

  if (embeddings) {
    const score = cosineSimilarity(embeddings.aVector, embeddings.bVector)
    return { status: score >= WORK_IDENTITY_MATCH_THRESHOLD ? 'match' : 'mismatch', score }
  }

  const setA = new Set(na.split(' '))
  const setB = new Set(nb.split(' '))
  let common = 0
  for (const t of setA) if (setB.has(t)) common++
  const union = setA.size + setB.size - common
  const overlap = union === 0 ? 0 : common / union
  return { status: overlap >= OVERLAP_THRESHOLD ? 'match' : 'mismatch', score: overlap }
}

/** A one-line, human-readable explanation of a work-name mismatch for a warning notice. */
export function workNameMismatchMessage(uploadedName: string, expectedName: string): string {
  return `The name of work in this file — "${uploadedName}" — is different from the work being processed here — "${expectedName}". Upload the details of the same work.`
}

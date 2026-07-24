import { cosineSimilarity } from './embeddingMatch'

export interface ColumnSpec {
  /** Human-readable name for this column's meaning — used for embedding similarity and error messages. */
  label: string
  /** Regex patterns tried first, in order (most specific first). Each must already be case-insensitive (`/i`). */
  patterns: RegExp[]
}

export interface ColumnEmbeddings {
  /** One embedding per header, same order/length as the `headers` array passed to resolveColumns. */
  headerVectors: number[][]
  /** One embedding per spec, same order/length as the `specs` array passed to resolveColumns. */
  labelVectors: number[][]
}

// Below this, an embedding-only match is treated as "no real match" — same
// caution as core/technicalSanction.ts's rate matching and
// core/createDocument.ts's placeholder matching: better to fail with a clear
// error than silently wire up the wrong column.
const EMBEDDING_THRESHOLD = 0.5

function normalizeHeader(h: string): string {
  return h.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase()
}

export interface ResolvedColumns {
  /** 0-based header index for each spec, keyed by ColumnSpec.label. */
  indexByLabel: Record<string, number>
  /** Labels of specs that only resolved via the embedding fallback, not a regex — worth a soft "please double-check" notice. */
  viaEmbedding: string[]
}

/**
 * Resolve every spec against `headers`, in order: try each spec's regex
 * patterns first (fast, exact/aliased matching — works offline with no
 * model, same as before this existed), and only when none match, fall back
 * to whichever unclaimed header is semantically closest to the spec's own
 * label (e.g. a column named "Approx Qty" or "Rate per Unit" that no fixed
 * regex anticipated) — the same graceful-degradation precedent as
 * core/technicalSanction.ts's rate matching and
 * core/createDocument.ts's placeholder matching. Omitting `embeddings`
 * reproduces regex-only matching exactly (a pure, synchronous fast path with
 * no model dependency).
 *
 * Throws naming the first spec that couldn't be resolved either way — every
 * spec must resolve to a *different* header (a header the embedding fallback
 * already picked for an earlier spec is never picked again).
 */
export function resolveColumns(
  headers: string[],
  specs: ColumnSpec[],
  embeddings?: ColumnEmbeddings
): ResolvedColumns {
  const normalized = headers.map(normalizeHeader)
  const claimed = new Set<number>()
  const indexByLabel: Record<string, number> = {}
  const viaEmbedding: string[] = []

  specs.forEach((spec, specIdx) => {
    let found = -1
    for (const re of spec.patterns) {
      const idx = normalized.findIndex((h, i) => !claimed.has(i) && re.test(h))
      if (idx !== -1) {
        found = idx
        break
      }
    }
    if (found === -1 && embeddings) {
      const labelVector = embeddings.labelVectors[specIdx]
      let bestIdx = -1
      let bestScore = -Infinity
      for (let i = 0; i < headers.length; i++) {
        if (claimed.has(i)) continue
        const score = cosineSimilarity(labelVector, embeddings.headerVectors[i])
        if (score > bestScore) {
          bestScore = score
          bestIdx = i
        }
      }
      if (bestIdx !== -1 && bestScore >= EMBEDDING_THRESHOLD) {
        found = bestIdx
        viaEmbedding.push(spec.label)
      }
    }
    if (found === -1) {
      throw new Error(`Could not find a "${spec.label}" column`)
    }
    indexByLabel[spec.label] = found
    claimed.add(found)
  })

  return { indexByLabel, viaEmbedding }
}

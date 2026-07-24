import { cosineSimilarity } from '@core/embeddingMatch'
import { BOQ_HEADERS } from '@core/boqHeaders'
import { WORKS_COLUMNS } from './worksSchema'

export type DocKind = 'works-list' | 'estimate' | 'boq'

interface KindProfile {
  kind: DocKind
  label: string
  /** Representative header/column names for this document kind. */
  headers: string[]
}

// A plain "detailed abstract estimate" sheet's own headers are rarely
// exactly this (see core/estimateExtract.ts's own regexes for the aliases
// actually matched against), but as a *prototype* for classification this
// captures the shape well enough to tell it apart from a Works List or BOQ.
const ESTIMATE_PROTOTYPE_HEADERS = ['S.No', 'Description', 'Quantity', 'Rate', 'Unit', 'Amount']

const PROFILES: KindProfile[] = [
  { kind: 'works-list', label: 'Works List', headers: WORKS_COLUMNS },
  { kind: 'estimate', label: 'Estimate', headers: ESTIMATE_PROTOTYPE_HEADERS },
  { kind: 'boq', label: 'BOQ', headers: BOQ_HEADERS }
]

const KIND_LABEL: Record<DocKind, string> = Object.fromEntries(PROFILES.map((p) => [p.kind, p.label])) as Record<
  DocKind,
  string
>

function tokenize(s: string): Set<string> {
  return new Set(
    s
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .trim()
      .split(/\s+/)
      .filter(Boolean)
  )
}

function overlapScore(a: Set<string>, b: Set<string>): number {
  if (a.size === 0 || b.size === 0) return 0
  let shared = 0
  for (const t of a) if (b.has(t)) shared++
  return shared / Math.max(a.size, b.size)
}

/** headers joined into one string, for embedding as a single "document" — see docKindPrototypeTexts(). */
export function joinedHeaderText(headers: string[]): string {
  return headers.join(', ')
}

/** The one descriptive text to embed for each known document kind, in the same order classifyHeaders scores them — pass these (plus the uploaded file's own joinedHeaderText) to api.embedTexts to build `DocEmbeddings`. */
export function docKindPrototypeTexts(): string[] {
  return PROFILES.map((p) => joinedHeaderText(p.headers))
}

export interface DocEmbeddings {
  /** Embedding of the uploaded file's own joinedHeaderText(headers). */
  documentVector: number[]
  /** One embedding per docKindPrototypeTexts() entry, same order. */
  profileVectors: number[][]
}

export interface ClassifyResult {
  kind: DocKind
  label: string
  score: number
}

/**
 * Guess which kind of document a set of uploaded headers belongs to, by
 * comparing them against each known kind's own column set — keyword overlap
 * of the header tokens (fast, offline, deterministic) combined with whole-
 * header-set semantic similarity (when `embeddings` is supplied, taking
 * `Math.max(keyword, semantic)` per kind) — the same hybrid approach used
 * elsewhere in the app (core/technicalSanction.ts, core/createDocument.ts).
 * Sorted best-first. This is advisory only, never a hard classification —
 * see mismatchHint, its only real consumer.
 */
export function classifyHeaders(headers: string[], embeddings?: DocEmbeddings): ClassifyResult[] {
  const uploadedTokens = new Set<string>()
  for (const h of headers) for (const t of tokenize(h)) uploadedTokens.add(t)

  return PROFILES.map((profile, i) => {
    const profileTokens = new Set<string>()
    for (const h of profile.headers) for (const t of tokenize(h)) profileTokens.add(t)
    const keyword = overlapScore(uploadedTokens, profileTokens)
    const semantic = embeddings ? cosineSimilarity(embeddings.documentVector, embeddings.profileVectors[i]) : 0
    return { kind: profile.kind, label: profile.label, score: Math.max(keyword, semantic) }
  }).sort((a, b) => b.score - a.score)
}

// The mismatch must be both reasonably confident and a clear enough margin
// over the expected kind — a close call (e.g. two kinds that legitimately
// share several column names) should stay silent rather than nag.
const MIN_CONFIDENCE = 0.3
const MIN_MARGIN = 0.1

/**
 * A short "this looks like a BOQ, not an Estimate" hint — only produced when
 * the top-scoring kind isn't `expectedKind` and clearly beats it. Intended to
 * be shown alongside (not instead of) a parsing failure, turning a generic
 * "could not find column X" error into an actionable one — never meant to
 * second-guess a file that parsed successfully.
 */
export function mismatchHint(headers: string[], expectedKind: DocKind, embeddings?: DocEmbeddings): string | null {
  const ranked = classifyHeaders(headers, embeddings)
  const top = ranked[0]
  if (!top || top.kind === expectedKind) return null
  const expectedScore = ranked.find((r) => r.kind === expectedKind)?.score ?? 0
  if (top.score < MIN_CONFIDENCE || top.score - expectedScore < MIN_MARGIN) return null
  return `This looks like a ${top.label}, not a${/^[aeiou]/i.test(KIND_LABEL[expectedKind]) ? 'n' : ''} ${KIND_LABEL[expectedKind]} — double-check you uploaded the right file.`
}

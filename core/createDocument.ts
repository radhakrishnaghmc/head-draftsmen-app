import { cosineSimilarity } from './embeddingMatch'

/** Shared with core/docx-edit.ts's OOXML-native placeholder find/fill — both formats must agree on what counts as a placeholder. */
export const PLACEHOLDER_RE = /\{\{\s*([^{}]+?)\s*\}\}/g

const NAMED_ENTITIES: Record<string, string> = {
  nbsp: ' ',
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'"
}

interface TextMap {
  /** Plain text with tags stripped and entities decoded. */
  text: string
  /** htmlIndex[i] = index into the original html string where text[i] came from. */
  htmlIndex: number[]
}

/**
 * Strip tags and decode entities while keeping a position map back to the
 * original HTML — so a {{Placeholder}} split across several tags is read
 * as one label, not one label per fragment plus markup noise. This matters
 * because real Word paste output very commonly splits what looks like one
 * run of text into several <span>/<b> elements (spell-check, language
 * tagging, bidi font sizing, etc.) even mid-placeholder, and matching
 * {{...}} directly against the raw HTML string picked up everything in
 * between — including the intervening tags themselves — as if it were the
 * label text, producing garbage like
 * "Amount&nbsp;</span></b><b style=\"...\">Estimate" that could never
 * resolve to a real column.
 */
function stripTagsWithMap(html: string): TextMap {
  let text = ''
  const htmlIndex: number[] = []
  let i = 0
  const n = html.length
  while (i < n) {
    const ch = html[i]
    if (ch === '<') {
      const close = html.indexOf('>', i)
      i = close === -1 ? n : close + 1
      continue
    }
    if (ch === '&') {
      const entity = /^&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/.exec(html.slice(i, i + 12))
      if (entity) {
        const raw = entity[1]
        const decoded =
          raw[0] === '#'
            ? String.fromCodePoint(parseInt(raw.slice(raw[1] === 'x' || raw[1] === 'X' ? 2 : 1), raw[1] === 'x' || raw[1] === 'X' ? 16 : 10))
            : NAMED_ENTITIES[raw.toLowerCase()]
        if (decoded !== undefined) {
          text += decoded
          htmlIndex.push(i)
          i += entity[0].length
          continue
        }
      }
    }
    text += ch
    htmlIndex.push(i)
    i++
  }
  return { text, htmlIndex }
}

interface PlaceholderOccurrence {
  label: string
  /** Index into the original html where this occurrence's "{{" begins. */
  htmlStart: number
  /** Index into the original html right after this occurrence's "}}" (exclusive). */
  htmlEnd: number
}

/** Every {{Label}} occurrence, tag-aware — safe to call with plain text or raw HTML. */
function findPlaceholderOccurrences(html: string): PlaceholderOccurrence[] {
  const { text, htmlIndex } = stripTagsWithMap(html)
  const occurrences: PlaceholderOccurrence[] = []
  PLACEHOLDER_RE.lastIndex = 0
  let match: RegExpExecArray | null
  while ((match = PLACEHOLDER_RE.exec(text)) !== null) {
    const startIdx = match.index
    const endIdx = match.index + match[0].length - 1
    if (startIdx >= htmlIndex.length || endIdx >= htmlIndex.length) continue
    // A run split across several tags/spans (Word does this constantly)
    // often leaves consecutive decoded &nbsp;s at the seams — collapse
    // those runs of whitespace so e.g. "Amount  of Estimate" (real gap: two
    // adjacent &nbsp; entities where two runs met) still matches the same
    // label as a normally-spaced "Amount of Estimate".
    const label = match[1].trim().replace(/\s+/g, ' ')
    occurrences.push({
      label,
      htmlStart: htmlIndex[startIdx],
      htmlEnd: htmlIndex[endIdx] + 1
    })
  }
  return occurrences
}

/** Every distinct {{Label}} found in the document, in first-seen order. Tag-aware — works the same whether given plain text or raw HTML. */
export function findPlaceholders(html: string): string[] {
  const seen = new Set<string>()
  for (const occ of findPlaceholderOccurrences(html)) seen.add(occ.label)
  return [...seen]
}

export interface PlaceholderMatch {
  label: string
  /** The Works List column this label resolved to, or null if nothing scored above the threshold. */
  column: string | null
  score: number
}

export interface EmbeddingVectors {
  labelVectors: number[][]
  columnVectors: number[][]
}

// A match below this score is treated as "no real match" — better to leave a
// placeholder unresolved (blanked) than confidently wire it to the wrong
// column, mirroring the same caution used for Give Technical Sanction's rate
// matching.
const MATCH_THRESHOLD = 0.5

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

// Dependency-free fallback for when the neural model is unavailable (offline
// failure, model not bundled, etc.) — plain word overlap, same graceful-
// degradation precedent as the rest of the app's embedding-backed matching.
function tokenOverlapScore(a: string, b: string): number {
  const ta = tokenize(a)
  const tb = tokenize(b)
  if (ta.size === 0 || tb.size === 0) return 0
  let shared = 0
  for (const t of ta) if (tb.has(t)) shared++
  return shared / Math.max(ta.size, tb.size)
}

/**
 * Match each placeholder label to the Works List column that means the same
 * thing — semantically (via the bundled neural embedding model) or by plain
 * word overlap, not just literal string equality. This is the "entirely new
 * logic": a pasted placeholder like {{Agency Name}} can resolve to a column
 * literally named "Name of the Contractor" as long as either signal scores
 * it close enough.
 *
 * When embeddings are supplied, a column's score is whichever is stronger,
 * `Math.max(tokenOverlapScore, embeddingScore)` — the same hybrid approach
 * core/technicalSanction.ts uses for rate matching, so a weak keyword overlap
 * doesn't drag down an otherwise-confident semantic match, and vice versa.
 * Omitting `embeddings` reproduces plain token-overlap matching exactly.
 */
export function matchPlaceholdersToColumns(
  labels: string[],
  columns: string[],
  embeddings?: EmbeddingVectors
): PlaceholderMatch[] {
  return labels.map((label, i) => {
    if (columns.length === 0) return { label, column: null, score: 0 }
    const scores = columns.map((column, j) => {
      const keyword = tokenOverlapScore(label, column)
      const semantic = embeddings ? cosineSimilarity(embeddings.labelVectors[i], embeddings.columnVectors[j]) : 0
      return Math.max(keyword, semantic)
    })
    let bestIdx = 0
    for (let j = 1; j < scores.length; j++) if (scores[j] > scores[bestIdx]) bestIdx = j
    const bestScore = scores[bestIdx]
    if (bestScore < MATCH_THRESHOLD) return { label, column: null, score: bestScore }
    return { label, column: columns[bestIdx], score: bestScore }
  })
}


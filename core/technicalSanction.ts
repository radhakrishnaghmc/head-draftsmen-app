import type { EstimateWorkItem } from './estimateExtract'
import type { RateEntry } from './rateDatabase'
import { normalizeUnit } from './rateDatabase'
import { buildRateIndexForMatching, findMatches } from './textMatch'
import { rankByEmbedding } from './embeddingMatch'

export interface MatchCandidate {
  description: string
  rate: string
  /** The Data Sheet's supporting material/labour/machinery buildup for this rate — carried through to the generated file's rate-analysis sheet. */
  breakdown: string[][]
  /** The matched Data Sheet entry's measuring unit, when known — surfaced so a wrong-unit match can be spotted. */
  unit?: string
  /** The matched Data Sheet entry's index code, when known (e.g. "RBR-EECD-8(c)"). */
  code?: string
}

// A departmental index code at the very start of an estimate item's
// description ("RBR-EECD-8(c): Earthwork excavation …", "BLD-CSTN-1-1 …"),
// with its optional sub-variant in parentheses. Anchored to the start so a
// clause/standard number sitting mid-sentence ("as per clause 305.1") is never
// mistaken for the item's own code.
const ITEM_CODE_RE = /^\s*([A-Za-z]{2,6}-[A-Za-z0-9.\-]*\d)\s*(?:\(\s*([A-Za-z]{1,4}|[ivx]{1,4})\s*\))?/i

/** The departmental code an estimate item cites, with its sub-variant when present (e.g. "RBR-EECD-8(c)"), or undefined when it cites none. */
export function extractItemCode(description: string): string | undefined {
  const m = ITEM_CODE_RE.exec(description)
  if (!m) return undefined
  return m[2] ? `${m[1]}(${m[2]})` : m[1]
}

// A concrete grade token ("M10", "M 30", "M-25") — only meaningful when the
// item is actually a concrete/mortar item, so it's read in that context alone
// (a stray "M20" in unrelated prose can't hijack the match).
const GRADE_IN_TEXT_RE = /\bM\s?-?\s?(\d{2,3})\b/i
const CONCRETE_CONTEXT_RE = /\b(concrete|rmc|rcc|pcc|design\s*mix|c\.?c\.?\s*pavement)\b/i

/** The concrete grade an estimate item specifies ("M10"), or undefined when it isn't a graded-concrete item. */
export function extractGrade(description: string): string | undefined {
  if (!CONCRETE_CONTEXT_RE.test(description)) return undefined
  const m = GRADE_IN_TEXT_RE.exec(description)
  return m ? `M${m[1]}` : undefined
}

/** Normalize a code for comparison — case- and punctuation-insensitive, so "RBR-EECD-8(c)" and "rbr eecd 8 c" match. */
function normCode(code: string): string {
  return code.toLowerCase().replace(/[^a-z0-9]/g, '')
}

function toCandidate(entry: RateEntry): MatchCandidate {
  return {
    description: entry.description,
    rate: entry.rate,
    breakdown: entry.breakdown,
    ...(entry.unit !== undefined ? { unit: entry.unit } : {}),
    ...(entry.code !== undefined ? { code: entry.code } : {})
  }
}

/** Deduplicate candidates by (description, rate) — the same rate duplicated across sheets isn't genuine ambiguity. */
function dedupeCandidates(entries: RateEntry[]): MatchCandidate[] {
  const seen = new Set<string>()
  const out: MatchCandidate[] = []
  for (const e of entries) {
    const key = `${e.description} ${e.rate}`
    if (seen.has(key)) continue
    seen.add(key)
    out.push(toCandidate(e))
  }
  return out
}

export interface ItemMatch {
  item: EstimateWorkItem
  /** Empty = unresolved (red). One = auto-resolved (green), unless semanticOnly. More than one = ambiguous — needs a user pick. */
  candidates: MatchCandidate[]
  bestScore: number
  /**
   * True when the winning match has no supporting keyword (TF-IDF) overlap
   * at all — it was found by embedding similarity alone. General-purpose
   * embedding models aren't fine-tuned on this domain's jargon and can
   * confidently conflate closely-related-but-distinct technical items (e.g.
   * a specific concrete grade vs. a generic concrete mix) — verified during
   * testing against a real estimate. So even a single semantic-only
   * candidate should be treated like an ambiguous match (needs a manual
   * pick), never auto-applied.
   */
  semanticOnly: boolean
}

// Cosine similarity below this is treated as "no real match" rather than a
// weak guess — better to flag red for manual entry than silently apply an
// unrelated rate to an official Technical Sanction document.
const MATCH_THRESHOLD = 0.35

// A single similarity match is only auto-filled (green) when it clears this
// higher bar; between MATCH_THRESHOLD and here it's shown for a human to
// confirm instead of being applied silently. Exact index-code and concrete-
// grade matches are scored 1 and so always clear it — the gate only holds back
// medium-confidence *similarity* guesses, which is exactly where a wrong rate
// slips onto an official Technical Sanction unnoticed: an item whose true rate
// isn't in the Data Sheet at all still draws some 0.4–0.65 lookalike, and that
// must be surfaced for review, never applied on its own. Verified end-to-end
// against a real estimate: genuine matches land at 0.95+, spurious lookalikes
// at 0.65 and below.
export const AUTO_APPLY_THRESHOLD = 0.7

export interface EmbeddingData {
  /** Sentence embedding for each item in `items`, same order, same length. */
  itemVectors: number[][]
  /** Sentence embedding for each entry in `rateIndex`, same order, same length. */
  entryVectors: number[][]
}

/**
 * Match each estimate item's description against the rate database. When the
 * best-matching description has several distinct rates attached to it (the
 * common "one paragraph, several manual/mechanical x depth-range rates"
 * pattern — see core/rateDatabase.ts), all of them are returned as
 * candidates for the caller to present for a manual pick, rather than
 * guessing which one applies.
 *
 * TF-IDF (core/textMatch.ts) only sees shared vocabulary, so a semantically
 * identical item worded very differently from its Data Sheet entry (a
 * paraphrase, a synonym, a different clause order) can score too low even
 * when a human would recognize it instantly. When `embeddings` is supplied
 * (precomputed sentence vectors from the bundled neural model — see
 * electron/embeddings.ts), each rate-database entry's score becomes
 * whichever signal is more confident: Math.max(tfidfScore, embeddingScore).
 * A strong match from either method is enough, so a weak lexical overlap
 * doesn't drag down a semantically obvious match or vice versa. Omitting
 * `embeddings` reproduces the original TF-IDF-only behavior exactly.
 */
export function matchEstimateItems(
  items: EstimateWorkItem[],
  rateIndex: RateEntry[],
  embeddings?: EmbeddingData
): ItemMatch[] {
  const idx = buildRateIndexForMatching(rateIndex)
  const positionOf = new Map<RateEntry, number>()
  rateIndex.forEach((e, i) => positionOf.set(e, i))

  // Index entries grouped by their normalized code, for exact code lookup.
  const byCode = new Map<string, RateEntry[]>()
  for (const e of rateIndex) {
    if (!e.code) continue
    const k = normCode(e.code)
    const arr = byCode.get(k)
    if (arr) arr.push(e)
    else byCode.set(k, [e])
  }

  return items.map((item, itemIdx) => {
    const itemUnit = normalizeUnit(item.unit)

    // 1. Exact index-code match — when the estimate cites a departmental code
    //    ("RBR-EECD-8(c): …"), the code pins the exact rate far more reliably
    //    than any wording similarity, so it wins outright. A code that resolves
    //    to a single rate auto-applies; one shared by several rates is surfaced
    //    for a pick, same as any ambiguous match.
    const itemCode = extractItemCode(item.description)
    if (itemCode) {
      const hits = byCode.get(normCode(itemCode))
      if (hits && hits.length > 0) {
        return { item, candidates: dedupeCandidates(hits), bestScore: 1, semanticOnly: false }
      }
    }

    // 2. Concrete-grade match — a "… M10 grade concrete …" item is pinned to the
    //    Data Sheet's rate for exactly that grade (its code carries the grade,
    //    e.g. "RBR-FNDN-4(M10)"), which a generic-concrete entry would otherwise
    //    outscore on shared boilerplate. Restricted to units compatible with the
    //    item's so a per-cum grade can't match a per-kg reinforcement line.
    const grade = extractGrade(item.description)
    if (grade) {
      const gradeKey = normCode(`(${grade})`)
      const gradeHits = rateIndex.filter(
        (e) => e.code && normCode(e.code).endsWith(gradeKey) && (!e.unit || !itemUnit || e.unit === itemUnit)
      )
      if (gradeHits.length > 0) {
        return { item, candidates: dedupeCandidates(gradeHits), bestScore: 1, semanticOnly: false }
      }
    }

    // 3. Fall back to lexical (TF-IDF) + semantic (embedding) similarity.
    const scoreByPosition = new Map<number, number>()
    const tfidfPositions = new Set<number>()
    for (const m of findMatches(item.description, idx)) {
      const pos = positionOf.get(m.entry)
      if (pos != null) {
        scoreByPosition.set(pos, m.score)
        tfidfPositions.add(pos)
      }
    }

    const itemVector = embeddings?.itemVectors[itemIdx]
    if (itemVector) {
      for (const { index: pos, score } of rankByEmbedding(itemVector, embeddings!.entryVectors)) {
        scoreByPosition.set(pos, Math.max(scoreByPosition.get(pos) ?? 0, score))
      }
    }

    let ranked = Array.from(scoreByPosition, ([pos, score]) => ({ pos, entry: rateIndex[pos], score })).sort(
      (a, b) => b.score - a.score
    )

    // Unit-aware filtering: drop candidates whose Data Sheet unit is known and
    // differs from the estimate item's own unit (a per-Sqm finishing rate can't
    // apply to a per-Each item). Entries with no recorded unit are kept (a
    // missing unit is "unknown", not "incompatible"), and the filter is skipped
    // when the item itself carries no usable unit, so it only ever removes
    // genuine mismatches, never leaves a resolvable item unresolved.
    if (itemUnit) {
      ranked = ranked.filter((m) => !m.entry.unit || m.entry.unit === itemUnit)
    }

    if (ranked.length === 0 || ranked[0].score < MATCH_THRESHOLD) {
      return { item, candidates: [], bestScore: ranked[0]?.score ?? 0, semanticOnly: false }
    }
    const topDescription = ranked[0].entry.description
    const semanticOnly = !tfidfPositions.has(ranked[0].pos)
    // dedupeCandidates collapses a (description, rate) pair duplicated across
    // sheets, so that repetition is not mistaken for genuine ambiguity.
    const candidates = dedupeCandidates(ranked.filter((m) => m.entry.description === topDescription).map((m) => m.entry))
    return { item, candidates, bestScore: ranked[0].score, semanticOnly }
  })
}

/**
 * Whether a match can be auto-filled (green) without a human check: exactly one
 * candidate, found by a trustworthy signal — an exact index-code / concrete-
 * grade match, or a strong similarity score (>= AUTO_APPLY_THRESHOLD) with real
 * keyword overlap (not embedding-only). Everything else — several candidate
 * rates, a semantic-only guess, or a medium-confidence lookalike — is surfaced
 * for the user to confirm or reject rather than applied silently.
 */
export function isAutoResolved(match: ItemMatch): boolean {
  return match.candidates.length === 1 && !match.semanticOnly && match.bestScore >= AUTO_APPLY_THRESHOLD
}

/** Whether a match has candidates but isn't confident enough to auto-fill — it needs a human pick/confirm. */
export function needsReview(match: ItemMatch): boolean {
  return match.candidates.length > 0 && !isAutoResolved(match)
}

export interface CellEdit {
  row: number
  col: number
  /** Which sheet to apply this edit to, for an estimate workbook that spreads its items across several sheets — undefined means the output file's default/single sheet. */
  sheetName?: string
  value?: string | number
  /**
   * 'green'/'red' recolor the description text (the Data Sheet/SoR rate
   * lookup above, a confidence signal on an otherwise-unedited description).
   * 'paleGreen'/'paleRed' fill the cell's background instead (Replicate
   * Technical Sanction below, a found/not-found signal on an exact
   * description match) — kept visually distinct so the two features' marks
   * are never confused for each other on the same sheet.
   */
  color?: 'green' | 'red' | 'paleGreen' | 'paleRed' | 'yellow'
}

export interface ResolvedItem {
  item: EstimateWorkItem
  /** The candidate the user (or auto-resolution) settled on, or null if left unresolved. */
  chosen: MatchCandidate | null
}

function numOrStr(s: string): number | string {
  const n = Number(String(s).replace(/,/g, '').trim())
  return Number.isFinite(n) && s.trim() !== '' ? Math.round(n * 100) / 100 : s
}

/**
 * Turn resolved matches into concrete cell edits for the output file. A
 * normal item gets both its rate and description cell updated (green). A
 * variant item (one of several measurement rows sharing one lead-row
 * description, e.g. a depth-range breakdown) only gets its rate and its own
 * row's color updated — its description text is left alone, since
 * overwriting it would clobber the shared lead-row wording for its siblings.
 * An unresolved item gets no rate change, just a red description cell.
 */
export function buildCellEdits(resolved: ResolvedItem[]): CellEdit[] {
  const edits: CellEdit[] = []
  for (const { item, chosen } of resolved) {
    if (item.descRow == null || item.descCol == null) continue
    const sheetName = item.sheetName
    if (chosen && item.rateRow != null && item.rateCol != null) {
      edits.push({ row: item.rateRow, col: item.rateCol, sheetName, value: numOrStr(chosen.rate) })
      edits.push(
        item.isVariant
          ? { row: item.descRow, col: item.descCol, sheetName, color: 'green' }
          : { row: item.descRow, col: item.descCol, sheetName, value: chosen.description, color: 'green' }
      )
    } else {
      edits.push({ row: item.descRow, col: item.descCol, sheetName, color: 'red' })
    }
  }
  return edits
}

/**
 * Build the backup "rate analysis" sheet content — the Data Sheet's own
 * material/labour/machinery buildup for each rate actually applied to the
 * estimate, one block per item, matching the convention seen in real
 * sanctioned estimates (a second sheet carrying the supporting calculations
 * behind each rate). Items left unresolved, or whose matched entry had no
 * captured breakdown, contribute nothing.
 */
export function buildRateAnalysisSheet(resolved: ResolvedItem[]): (string | number)[][] {
  const rows: (string | number)[][] = []
  for (const { chosen } of resolved) {
    if (!chosen || chosen.breakdown.length === 0) continue
    rows.push([`Rate Analysis: ${chosen.description}`])
    for (const r of chosen.breakdown) rows.push(r)
    rows.push([])
  }
  return rows
}

/**
 * Replicate Technical Sanction: a separate, simpler sibling to the Data
 * Sheet/SoR rate lookup above — instead of retrieving a rate by keyword/
 * semantic search, it reuses a rate already sanctioned on a *prior*
 * estimate, for an item whose description is the same work repeated (e.g.
 * this year's estimate re-quoting last year's already-sanctioned items).
 * Exact description match only, no scoring/ranking — either a prior
 * estimate carried that exact wording, or it didn't.
 */
export interface ReplicateMatch {
  item: EstimateWorkItem
  /** The rate carried over from a matching already-sanctioned item, or undefined when no source estimate had that description. */
  rate?: string
}

/** Collapse a description to a bare comparison key — trimmed, whitespace-collapsed, case-insensitive, trailing-punctuation-stripped — so the same wording printed with different spacing/casing/a stray trailing full stop across two estimate files still counts as the same item. */
export function normalizeItemDescription(description: string): string {
  return description
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[.,;]+$/, '')
}

/**
 * Build a description -> rate index from one or more already-sanctioned
 * source estimates' items (merged across every file uploaded, by the
 * caller re-calling this with the accumulated item list). The first rate
 * seen for a given description wins — if two source estimates disagree
 * on an item's rate, the first upload stays authoritative rather than
 * silently overridden by a later one.
 */
export function buildReplicateIndex(oldItems: EstimateWorkItem[]): Map<string, string> {
  const index = new Map<string, string>()
  for (const item of oldItems) {
    if (!item.description || !item.rate) continue
    const key = normalizeItemDescription(item.description)
    if (!index.has(key)) index.set(key, item.rate)
  }
  return index
}

/** Match each new estimate item's description, verbatim, against the already-sanctioned source index. */
export function matchByExactDescription(newItems: EstimateWorkItem[], index: Map<string, string>): ReplicateMatch[] {
  return newItems.map((item) => {
    const rate = index.get(normalizeItemDescription(item.description))
    return rate !== undefined ? { item, rate } : { item }
  })
}

/**
 * Turn exact matches into cell edits for the output file. A found item's
 * rate cell is filled and marked yellow; an unmatched item gets no rate
 * change, just its rate cell marked pale red — see CellEdit's color field
 * for why this uses the pale fill colors instead of the Data Sheet lookup's
 * green/red text colors.
 */
export function buildReplicateCellEdits(matches: ReplicateMatch[]): CellEdit[] {
  const edits: CellEdit[] = []
  for (const { item, rate } of matches) {
    if (item.rateRow != null && rate !== undefined && item.rateCol != null) {
      edits.push({ row: item.rateRow, col: item.rateCol, sheetName: item.sheetName, value: numOrStr(rate), color: 'yellow' })
    } else if (item.rateRow != null && item.rateCol != null) {
      edits.push({ row: item.rateRow, col: item.rateCol, sheetName: item.sheetName, color: 'paleRed' })
    }
  }
  return edits
}

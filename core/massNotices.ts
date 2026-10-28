/**
 * "Generate Mass Notices" — pure logic behind Settings' bulk workspace that
 * turns a folder of tender documents into "Notice for Non-conclusion of
 * Agreement" tiles. The workspace (src/components/MassNoticesWorkspace.tsx)
 * does the file reading/OCR; everything decidable from the parsed results
 * lives here so it stays unit-testable:
 *
 *  - pairing each L-1 selection sheet with the LOA (Online Intimation) of the
 *    SAME work — same agency/NIT rule the Intimation and Work Order tabs use
 *    (core/sameWorkCheck.ts); an L-1 with no LOA gets no notice;
 *  - reading the Monitoring Format's "Agreement" column: a real date means the
 *    agreement is concluded (no notice); blank or words ("Not yet", "Under
 *    process") means it isn't (notice).
 */
import type { IntimationNotice } from './intimationNotice'
import type { TenderEvaluation } from './tenderEvaluationPdf'
import type { MonitoringFormatWorkRow } from './monitoringFormat'
import { checkSameWork } from './sameWorkCheck'
import { normalizeWorkNameForMatch } from './tenderAgents/nameOfWork'
import { rankByEmbedding, workNamesConflict, WORK_IDENTITY_MATCH_THRESHOLD } from './embeddingMatch'

/** An L-1 selection sheet and the LOA found for the same work. */
export interface L1LoaPair {
  l1: TenderEvaluation
  loa: IntimationNotice
  l1File: string
  /** File the LOA (or, for via 'supplier', the Supplier Details page) was read from. */
  loaFile: string
  /** 'supplier' = no LOA in the folder; the L-1 was confirmed against the bidder's Supplier Details page instead. */
  via: 'loa' | 'supplier'
}

export interface PairingResult {
  pairs: L1LoaPair[]
  /** L-1 sheets for which no LOA of the same work was in the folder — no notice is generated for these. */
  noLoa: { l1: TenderEvaluation; l1File: string }[]
}

/** The eGP portal's "Supplier Details" page — the bidder's profile for one tender (carries the portal Tender ID, company name, address, phone). */
export type SupplierDetails = IntimationNotice & { tenderId: string }

export interface ParsedDoc<T> {
  file: string
  doc: T
}

function near(a: number | undefined, b: number | undefined): boolean {
  return a != null && b != null && Math.abs(a - b) <= 1
}

/**
 * Pairs every L-1 sheet with the LOA that belongs to the same work (agency
 * and NIT No agree — checkSameWork). A status of 'unknown' (nothing to
 * compare) is NOT treated as a pairing: with no way to verify, generating a
 * notice could send one work's agency another work's notice. When several
 * LOAs fit one L-1 (one agency winning several items of a NIT) the one whose
 * contract value agrees wins, and each LOA is used for at most one L-1. An L-1
 * with no LOA at all is accepted on its Supplier Details page when the Tender
 * ID is the same and the page's company is the L-1 agency.
 */
export function pairL1WithLoa(
  l1s: ParsedDoc<TenderEvaluation>[],
  loas: ParsedDoc<IntimationNotice>[],
  suppliers: ParsedDoc<SupplierDetails>[] = []
): PairingResult {
  const used = new Set<number>()
  const pairs: L1LoaPair[] = []
  const noLoa: PairingResult['noLoa'] = []
  for (const { doc: l1, file: l1File } of l1s) {
    const fits = loas
      .map((l, i) => ({ l, i }))
      .filter(({ l, i }) => !used.has(i) && checkSameWork(l.doc, l1).status === 'match')
    const best = fits.find(({ l }) => near(l.doc.contractRupees, l1.contractRupees)) ?? fits[0]
    if (!best) {
      // No LOA: fall back to the Supplier Details page — same portal Tender ID
      // AND the L-1 agency is the page's company (same agency rule as above).
      const sup = l1.tenderId ? suppliers.find((x) => x.doc.tenderId === l1.tenderId && checkSameWork({ agencyName: x.doc.agencyName }, l1).status === 'match') : undefined
      if (sup) {
        pairs.push({ l1, loa: { ...sup.doc, nitNo: l1.noticeNo, nitDate: l1.noticeDate }, l1File, loaFile: sup.file, via: 'supplier' })
      } else {
        noLoa.push({ l1, l1File })
      }
      continue
    }
    used.add(best.i)
    pairs.push({ l1, loa: best.l.doc, l1File, loaFile: best.l.file, via: 'loa' })
  }
  return { pairs, noLoa }
}

const MONTH = '(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*'
const DATE_PATTERNS: RegExp[] = [
  /\b\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{2,4}\b/, // 12.03.2026, 12/3/26, 12-03-2026
  /\b\d{4}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{1,2}\b/, // 2026-03-12
  new RegExp(`\\b\\d{1,2}(?:st|nd|rd|th)?[\\s./-]*${MONTH}[\\s.,/-]*\\d{2,4}\\b`, 'i'), // 12-Mar-2026, 12th March 2026
  new RegExp(`\\b${MONTH}[\\s.,/-]*\\d{1,2}(?:st|nd|rd|th)?[\\s.,/-]+\\d{4}\\b`, 'i') // March 12, 2026
]
/** Wording that says the agreement is NOT done even when a date appears in the same cell (e.g. "Not yet, LOA dt 12.03.2026"). */
const PENDING_WORDS =
  /\b(?:not|yet|pending|process|awaiting|await|nil|n\/?a|tbd|to be|will|under|after|before|soon|na)\b/i

/**
 * True when a Monitoring Format "Agreement" cell holds an actual agreement
 * date. Blank cells, and cells with words instead of a date ("Not yet",
 * "Under process", "LOA issued"), are NOT dates — the agreement hasn't been
 * concluded, so a notice is due.
 */
export function hasAgreementDate(cell: string | undefined): boolean {
  const s = (cell ?? '').replace(/\s+/g, ' ').trim()
  if (!s || PENDING_WORDS.test(s)) return false
  if (DATE_PATTERNS.some((re) => re.test(s))) return true
  // A date cell exported as a bare Excel serial number (e.g. 46093).
  const serial = /^\d{5}(?:\.\d+)?$/.exec(s)
  return !!serial && Number(s) > 36000 && Number(s) < 60000
}

export interface MfWorkMatch {
  row: MonitoringFormatWorkRow
  /** Matched by similarity rather than an exact normalized name. */
  fuzzy: boolean
}

/**
 * Finds the Monitoring Format row for a tender's Name of Work — exact
 * normalized name first (same normalization the Works List match uses), then,
 * when the caller supplies embedding vectors, the closest row above the
 * work-identity threshold; without vectors, a strict word-overlap fallback.
 * Rows whose names describe visibly different works (workNamesConflict) are
 * never accepted.
 */
export function findMonitoringWork(
  rows: MonitoringFormatWorkRow[],
  workName: string | undefined,
  embeddings?: { workNameVector: number[]; rowNameVectors: number[][] }
): MfWorkMatch | undefined {
  const target = normalizeWorkNameForMatch(workName)
  if (!target) return undefined
  const exact = rows.find((r) => normalizeWorkNameForMatch(r.workName) === target)
  if (exact) return { row: exact, fuzzy: false }

  if (embeddings) {
    const [best] = rankByEmbedding(embeddings.workNameVector, embeddings.rowNameVectors)
    if (
      best &&
      best.score >= WORK_IDENTITY_MATCH_THRESHOLD &&
      !workNamesConflict(workName ?? '', rows[best.index].workName)
    ) {
      return { row: rows[best.index], fuzzy: true }
    }
    return undefined
  }

  const tokens = new Set(target.split(' ').filter(Boolean))
  let bestRow: MonitoringFormatWorkRow | undefined
  let bestScore = 0
  for (const r of rows) {
    const other = new Set(normalizeWorkNameForMatch(r.workName).split(' ').filter(Boolean))
    let common = 0
    for (const t of tokens) if (other.has(t)) common++
    const union = tokens.size + other.size - common
    const score = union === 0 ? 0 : common / union
    if (score > bestScore) {
      bestScore = score
      bestRow = r
    }
  }
  return bestRow && bestScore >= 0.8 && !workNamesConflict(workName ?? '', bestRow.workName)
    ? { row: bestRow, fuzzy: true }
    : undefined
}

/** Why one tender did or didn't get a notice — what the workspace's result table shows. */
export type MassNoticeDecision = { kind: 'generate'; mfNote: string } | { kind: 'skip'; reason: string }

/**
 * The Monitoring Format rule for one paired tender: skip when its work's
 * Agreement column holds a date; otherwise (blank, words, or the work isn't in
 * the sheet at all) generate. `mfRows` undefined/empty = no Monitoring Format
 * has been imported for this office, so nothing can be ruled out.
 */
export function decideFromMonitoringFormat(
  match: MfWorkMatch | undefined,
  mfRows: MonitoringFormatWorkRow[] | undefined
): MassNoticeDecision {
  if (!mfRows || mfRows.length === 0)
    return { kind: 'generate', mfNote: 'No Monitoring Format imported — agreement date not checked' }
  if (!match) return { kind: 'generate', mfNote: 'Work not available in Monitoring Format' }
  const cell = match.row.agreementDate.trim()
  if (hasAgreementDate(cell)) return { kind: 'skip', reason: `Agreement already concluded (${cell})` }
  return {
    kind: 'generate',
    mfNote: cell ? `Agreement column says “${cell}” — not a date` : 'Agreement date blank in Monitoring Format'
  }
}

function nk(s: string | undefined): string {
  return (s ?? '').toLowerCase().replace(/[^a-z0-9]+/g, '')
}

/**
 * Duplicate notices — the same tender noticed twice. Identity is the portal
 * Tender ID when the notice has one; older notices (made before it was
 * stored) fall back to NIT No + agency + work name. Of each duplicate set the
 * one to KEEP is, in order: a resolved notice, one the user edited, the
 * earliest. Returns the ids of the rest. Already-deleted notices are ignored.
 */
export function duplicateNoticeIds(
  items: { id: string; tenderId?: string; nitNo?: string; agencyName?: string; workName?: string; done: boolean; edits?: Record<string, string>; deletedAt?: string }[]
): string[] {
  const sets = new Map<string, typeof items>()
  for (const it of items) {
    if (it.deletedAt) continue
    const key = it.tenderId ? `t|${it.tenderId.trim()}` : `n|${nk(it.nitNo)}|${nk(it.agencyName)}|${nk(it.workName)}`
    sets.set(key, [...(sets.get(key) ?? []), it])
  }
  const out: string[] = []
  for (const list of sets.values()) {
    if (list.length < 2) continue
    const rank = (it: (typeof items)[number]) => (it.done ? 2 : 0) + (it.edits && Object.keys(it.edits).length ? 1 : 0)
    let keep = list[0]
    for (const it of list) if (rank(it) > rank(keep)) keep = it
    for (const it of list) if (it !== keep) out.push(it.id)
  }
  return out
}

import type { ExcelTable } from './types'
import type { TenderEvaluation } from './tenderEvaluationPdf'
import type { IntimationNotice } from './intimationNotice'
import { rankByEmbedding } from './embeddingMatch'
import { rupeesToCell } from './worksAmounts'

// A match below this score is treated as "no real match" — same threshold
// used for work-name matching elsewhere (core/tenderMatch.ts, core/scheduleA.ts).
const EMBEDDING_THRESHOLD = 0.5

const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')

export interface WorksTenderUpdateResult {
  table: ExcelTable
  /** Evaluations whose Name of Work matched a Works List row and updated it. */
  matchedCount: number
  /** Names of Work from the PDFs that matched no Works List row, for a "couldn't place these" notice. */
  unmatched: string[]
  /**
   * Works List row indices that were matched & updated, in the order the
   * matches were made. Lets a caller select the matched row even when the
   * match came from the embedding fallback (wording drift) — where there's no
   * exact name to re-derive the index from, so the caller can't find it itself.
   */
  matchedRowIndices: number[]
}

/**
 * Every Works List column derivable from an award, and how to render each — from
 * the L-1 evaluation (`ev`) and, when the Online Intimation for the same work is
 * also to hand (`notice`), the fields only it carries (the agency's postal
 * address) plus a fallback for the agency name / contract value when the L-1
 * sheet didn't spell them out. EMD @ 1%/1.5% and ASD are computed from the ECV
 * and Tender % here so the database shows them straight away, matching how
 * core/worksAmounts.ts derives them for the documents. Only fields a source
 * actually carried are written; blanks never overwrite an existing value.
 */
function valuesFor(ev: TenderEvaluation, notice?: IntimationNotice): Record<string, string> {
  const out: Record<string, string> = {}
  if (ev.tenderId) out['Tender ID'] = ev.tenderId
  if (ev.noticeNo) out['Tender Notice No'] = ev.noticeNo
  if (ev.noticeDate) out['Tender notice Date'] = ev.noticeDate
  const ecv = ev.ecvRupees ?? notice?.ecvRupees
  if (ecv != null) {
    out['ECV'] = rupeesToCell(ecv)
    // EMD @ 1% / 1.5% of ECV, and ASD = (Tender% - 25%) of ECV once past 25%
    // (0 below it) — the same rules core/worksAmounts.ts applies for documents.
    out['EMD 1%'] = rupeesToCell(Math.round(ecv * 0.01))
    out['EMD 1.5%'] = rupeesToCell(Math.round(ecv * 0.015))
    if (ev.tenderPercentage != null) {
      out['ASD'] = rupeesToCell(ev.tenderPercentage > 25 ? Math.round(ecv * ((ev.tenderPercentage - 25) / 100)) : 0)
    }
  }
  const agency = ev.l1AgencyName || notice?.agencyName
  if (agency) out['Name of the Agency'] = agency
  if (notice?.address) out['Address of the agency'] = notice.address
  if (ev.tenderPercentage != null) out['Tender Percentage'] = String(ev.tenderPercentage)
  const contract = ev.contractRupees ?? notice?.contractRupees
  if (contract != null) out['Contract Amount'] = rupeesToCell(contract)
  return out
}

/**
 * Update the Works List from one or more parsed tender-evaluation PDFs (see
 * core/tenderEvaluationPdf.ts), matching each PDF's "Name of Work" to a
 * Works List row's "Name of the work" — exact normalized match first, then
 * an embedding fallback for the wording drift between a tender's title and
 * the Works List's own entry (abbreviations, punctuation) — and writing that
 * row's Tender ID, Tender Notice No, Tender notice Date, ECV, Name of the
 * Agency (the L-1 bidder), Tender Percentage, and Contract Amount. Only fields the PDF
 * actually carried are written; existing values for those fields are
 * overwritten (the PDF is the authoritative post-award record). A PDF whose
 * work name matches nothing is reported in `unmatched` rather than guessed at.
 *
 * `embeddings` (optional) supplies the caller-computed vectors for every
 * Works List row name and every PDF work name, in the same order as
 * `table.rows` and `evaluations` respectively.
 */
export function updateWorksListFromEvaluations(
  table: ExcelTable,
  evaluations: TenderEvaluation[],
  embeddings?: { rowNameVectors: number[][]; evalNameVectors: number[][] },
  // The Online Intimation for the same work, when available (the Work Order /
  // Agreement flow uploads both) — supplies the agency's postal address and a
  // fallback for agency name / contract value. A batch import (many
  // evaluations, no single intimation) omits it.
  notice?: IntimationNotice
): WorksTenderUpdateResult {
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  const usable = evaluations.filter((e) => (e.nameOfWork ?? '').trim())
  if (!nameHeader || usable.length === 0) {
    return { table, matchedCount: 0, unmatched: usable.map((e) => e.nameOfWork!).filter(Boolean), matchedRowIndices: [] }
  }

  const rowKeyToIndex = new Map<string, number>()
  table.rows.forEach((row, i) => {
    const key = norm(row[nameHeader] ?? '')
    if (key && !rowKeyToIndex.has(key)) rowKeyToIndex.set(key, i)
  })

  const rows = table.rows.map((r) => ({ ...r }))
  let matchedCount = 0
  const unmatched: string[] = []
  const matchedRowIndices: number[] = []

  usable.forEach((ev, evIndex) => {
    const target = norm(ev.nameOfWork!)
    let idx = rowKeyToIndex.get(target)
    if (idx == null && embeddings) {
      const [best] = rankByEmbedding(embeddings.evalNameVectors[evIndex], embeddings.rowNameVectors)
      if (best && best.score >= EMBEDDING_THRESHOLD) idx = best.index
    }
    if (idx == null) {
      unmatched.push(ev.nameOfWork!)
      return
    }
    matchedCount++
    matchedRowIndices.push(idx)
    // The single-evaluation Work Order flow passes the matching intimation; a
    // batch import (many evaluations) has none, so its rows fill from the L-1
    // sheet alone.
    Object.assign(rows[idx], valuesFor(ev, evaluations.length === 1 ? notice : undefined))
  })

  return { table: { ...table, rows }, matchedCount, unmatched, matchedRowIndices }
}

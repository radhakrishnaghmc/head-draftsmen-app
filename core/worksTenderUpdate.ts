import type { ExcelTable } from './types'
import type { TenderEvaluation } from './tenderEvaluationPdf'
import { rankByEmbedding } from './embeddingMatch'
import { rupeesToLakhsString } from './worksAmounts'

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
}

/** The Works List columns a tender evaluation fills, and how to render each field. */
function valuesFor(ev: TenderEvaluation): Record<string, string> {
  const out: Record<string, string> = {}
  if (ev.tenderId) out['Tender ID'] = ev.tenderId
  if (ev.noticeNo) out['Tender Notice No'] = ev.noticeNo
  if (ev.noticeDate) out['Tender notice Date'] = ev.noticeDate
  if (ev.ecvRupees != null) out['ECV'] = rupeesToLakhsString(ev.ecvRupees)
  if (ev.l1AgencyName) out['Name of the Agency'] = ev.l1AgencyName
  if (ev.tenderPercentage != null) out['Tender Percentage'] = String(ev.tenderPercentage)
  if (ev.contractRupees != null) out['Contract Amount'] = rupeesToLakhsString(ev.contractRupees)
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
  embeddings?: { rowNameVectors: number[][]; evalNameVectors: number[][] }
): WorksTenderUpdateResult {
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  const usable = evaluations.filter((e) => (e.nameOfWork ?? '').trim())
  if (!nameHeader || usable.length === 0) {
    return { table, matchedCount: 0, unmatched: usable.map((e) => e.nameOfWork!).filter(Boolean) }
  }

  const rowKeyToIndex = new Map<string, number>()
  table.rows.forEach((row, i) => {
    const key = norm(row[nameHeader] ?? '')
    if (key && !rowKeyToIndex.has(key)) rowKeyToIndex.set(key, i)
  })

  const rows = table.rows.map((r) => ({ ...r }))
  let matchedCount = 0
  const unmatched: string[] = []

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
    Object.assign(rows[idx], valuesFor(ev))
  })

  return { table: { ...table, rows }, matchedCount, unmatched }
}

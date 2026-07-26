import type { ExcelTable } from './types'
import type { TenderEvaluation } from './tenderEvaluationPdf'
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
}

/** The Works List columns a tender evaluation fills, and how to render each field. */
function valuesFor(ev: TenderEvaluation): Record<string, string> {
  const out: Record<string, string> = {}
  if (ev.tenderId) out['Tender ID'] = ev.tenderId
  if (ev.noticeNo) out['Tender Notice No'] = ev.noticeNo
  if (ev.noticeDate) out['Tender notice Date'] = ev.noticeDate
  if (ev.ecvRupees != null) out['ECV'] = rupeesToCell(ev.ecvRupees)
  if (ev.l1AgencyName) out['Name of the Agency'] = ev.l1AgencyName
  if (ev.tenderPercentage != null) out['Tender Percentage'] = String(ev.tenderPercentage)
  if (ev.contractRupees != null) out['Contract Amount'] = rupeesToCell(ev.contractRupees)
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

export interface AgencyAddressResult {
  table: ExcelTable
  /** Rows whose "Address of the agency" was set from the per-agency address map. */
  filledCount: number
}

/**
 * Fill each Works List row's "Address of the agency" from a per-agency
 * address map (keyed by normalized "Name of the Agency"), harvested from the
 * intimation notices found alongside the tender-evaluation PDFs. Matched by
 * agency name alone — never by Name of the work — because one agency has
 * exactly one address, so once a work's agency is known the same address
 * applies to every work that agency won. Only rows that already carry a
 * "Name of the Agency" are considered; the address is written when it differs
 * from what's there (same agency = same address, so this is idempotent and
 * authoritative).
 */
export function applyAgencyAddresses(
  table: ExcelTable,
  addressByAgency: Map<string, string>
): AgencyAddressResult {
  const agencyHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the agency')
  const addressHeader = table.headers.find((h) => h.trim().toLowerCase() === 'address of the agency')
  if (!agencyHeader || !addressHeader || addressByAgency.size === 0) return { table, filledCount: 0 }

  let filledCount = 0
  const rows = table.rows.map((r) => {
    const agency = norm(r[agencyHeader] ?? '')
    if (!agency) return r
    const address = addressByAgency.get(agency)
    if (!address || (r[addressHeader] ?? '') === address) return r
    filledCount++
    return { ...r, [addressHeader]: address }
  })
  return { table: { ...table, rows }, filledCount }
}

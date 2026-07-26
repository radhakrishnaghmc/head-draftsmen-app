import type { ExcelTable } from './types'
import { rankByEmbedding } from './embeddingMatch'
import { rupeesToCell } from './worksAmounts'

// A match below this score is treated as "no real match" — same threshold
// used for work-name matching elsewhere (core/worksAmounts.ts,
// core/scheduleA.ts).
const EMBEDDING_THRESHOLD = 0.5

// Raw tender row column indices — the Telangana e-procurement portal's own
// DataTables response (see electron/main.ts's buildTenderUrl/fetchTenders),
// matching src/components/SearchTender.tsx's own COLUMNS mapping exactly.
const TENDER_ID_COL = 1
const NOTICE_NO_COL = 2
const WORK_NAME_COL = 4
const ECV_COL = 5

export interface TenderRecord {
  tenderId: string
  noticeNo: string
  workName: string
  /** The portal's own ECV figure, in rupees (its listing is not Lakhs-denominated, unlike this app's Works List). */
  ecvRupees: number
}

/** Parses raw tender-portal rows into the fields this app cares about, dropping any row with no usable Tender ID or Work name to match against. */
export function parseTenderRows(rows: string[][]): TenderRecord[] {
  return rows
    .map((r) => ({
      tenderId: (r[TENDER_ID_COL] ?? '').trim(),
      noticeNo: (r[NOTICE_NO_COL] ?? '').trim(),
      workName: (r[WORK_NAME_COL] ?? '').trim(),
      ecvRupees: Number(r[ECV_COL]) || 0
    }))
    .filter((t) => t.tenderId && t.workName)
}

export interface FillFromTendersResult {
  table: ExcelTable
  matchedCount: number
}

/**
 * Match each Works List row's "Name of the work" against `tenders` (exact
 * normalized match, embedding fallback) and fill that row's ECV/Tender
 * Notice No/Tender ID from whichever tender matches best — overwriting
 * any existing value, since this is meant to refresh against the tender
 * portal's own live, authoritative data rather than merely fill gaps. A row
 * with no usable "Name of the work", or no tender matching above the
 * threshold, is left untouched.
 *
 * When more than one tender shares the same (normalized) work name — a
 * corrigendum/re-tender, or two unrelated works happening to be worded the
 * same — the first one seen wins for the exact-match path; embeddings, when
 * supplied, are only consulted for a row the exact pass didn't resolve.
 */
export function fillWorksListFromTenders(
  table: ExcelTable,
  tenders: TenderRecord[],
  embeddings?: { rowNameVectors: number[][]; tenderNameVectors: number[][] }
): FillFromTendersResult {
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  if (!nameHeader || tenders.length === 0) return { table, matchedCount: 0 }

  const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')
  const tendersByName = new Map<string, TenderRecord>()
  for (const t of tenders) {
    const key = norm(t.workName)
    if (!tendersByName.has(key)) tendersByName.set(key, t)
  }

  let matchedCount = 0
  const rows = table.rows.map((row, i) => {
    const target = norm(row[nameHeader] ?? '')
    if (!target) return row

    let match = tendersByName.get(target)
    if (!match && embeddings) {
      const [best] = rankByEmbedding(embeddings.rowNameVectors[i], embeddings.tenderNameVectors)
      if (best && best.score >= EMBEDDING_THRESHOLD) match = tenders[best.index]
    }
    if (!match) return row

    matchedCount++
    return {
      ...row,
      ECV: rupeesToCell(match.ecvRupees),
      'Tender Notice No': match.noticeNo,
      'Tender ID': match.tenderId
    }
  })

  return { table: { ...table, rows }, matchedCount }
}

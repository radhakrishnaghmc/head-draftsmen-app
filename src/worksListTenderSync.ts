import { api } from './ipc'
import type { ExcelTable } from '@core/types'
import { parseTenderRows, fillWorksListFromTenders } from '@core/tenderMatch'
import type { TenderRecord } from '@core/tenderMatch'

// Large enough to keep the number of round trips down for a circle with many
// tenders, without asking the portal for an unreasonably large page.
const PAGE_SIZE = 100

/**
 * Fetches every tender matching `search` across as many pages as the portal
 * reports (or, when it doesn't report a reliable total, until a page comes
 * back short) — sequentially, one request at a time. The portal is reached
 * through a single shared hidden browser window (electron/main.ts's
 * fetchTenders/tenderWindow), so concurrent searches would have it
 * navigating to two URLs at once instead of actually running in parallel.
 */
async function fetchAllTenders(search: string): Promise<TenderRecord[]> {
  const all: TenderRecord[] = []
  let start = 0
  for (;;) {
    const res = await api.searchTenders({ start, length: PAGE_SIZE, type: 'current', search })
    all.push(...parseTenderRows(res.data))
    start += PAGE_SIZE
    const hasMore = res.total >= 0 ? start < res.total : res.data.length === PAGE_SIZE
    if (!hasMore || res.data.length === 0) break
  }
  return all
}

export interface TenderSyncProgress {
  circleIndex: number
  circleCount: number
  circle: string
}

export interface TenderSyncResult {
  table: ExcelTable
  matchedCount: number
  tenderCount: number
  circleCount: number
}

/**
 * For every distinct Circle already on the Works List, searches the tender
 * portal by that circle's own name — there's no dedicated "by circle"
 * filter on the portal, so this reuses the same keyword search Search
 * Tender itself sends (Tender ID / IFB No / Name of Work) — gathers every
 * result, and fills each Works List row's ECV/Tender Notice No/Tender ID
 * from whichever tender's own Work name matches best (core/tenderMatch.ts).
 * Existing values are overwritten: the point is to refresh against the
 * portal's own live data, not merely fill gaps.
 */
export async function syncWorksListFromTenderPortal(
  table: ExcelTable,
  onProgress?: (p: TenderSyncProgress) => void
): Promise<TenderSyncResult> {
  const circleHeader = table.headers.find((h) => h.trim().toLowerCase() === 'circle')
  const circles = circleHeader
    ? Array.from(new Set(table.rows.map((r) => (r[circleHeader] ?? '').trim()).filter(Boolean)))
    : []

  const allTenders: TenderRecord[] = []
  const seenTenderIds = new Set<string>()
  for (let i = 0; i < circles.length; i++) {
    onProgress?.({ circleIndex: i, circleCount: circles.length, circle: circles[i] })
    const tenders = await fetchAllTenders(circles[i])
    for (const t of tenders) {
      if (seenTenderIds.has(t.tenderId)) continue
      seenTenderIds.add(t.tenderId)
      allTenders.push(t)
    }
  }

  let embeddings: { rowNameVectors: number[][]; tenderNameVectors: number[][] } | undefined
  const nameHeader = table.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  if (nameHeader && allTenders.length > 0) {
    try {
      const rowNames = table.rows.map((r) => r[nameHeader] ?? '')
      const vectors = await api.embedTexts([...rowNames, ...allTenders.map((t) => t.workName)])
      embeddings = {
        rowNameVectors: vectors.slice(0, rowNames.length),
        tenderNameVectors: vectors.slice(rowNames.length)
      }
    } catch {
      // Semantic matching unavailable — exact-name matching alone still applies.
    }
  }

  const { table: updated, matchedCount } = fillWorksListFromTenders(table, allTenders, embeddings)
  return { table: updated, matchedCount, tenderCount: allTenders.length, circleCount: circles.length }
}

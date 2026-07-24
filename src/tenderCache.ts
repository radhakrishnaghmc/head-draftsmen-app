import { api } from './ipc'
import type { TenderQuery, TenderResult } from '@core/types'

/**
 * A tiny in-memory cache for tender lookups, keyed by the query. It lets us
 * prefetch the default Search Tender view in the background as soon as the app
 * opens, so switching to that tab shows results instantly instead of firing a
 * fresh network request each time the component mounts.
 */
const cache = new Map<string, Promise<TenderResult>>()

function keyOf(q: TenderQuery): string {
  return `${q.type}|${q.length}|${q.start}|${q.search}`
}

/** The query warmed on startup and used for the tab's default view. */
export const DEFAULT_TENDER_QUERY: TenderQuery = {
  start: 0,
  length: 25,
  type: 'current',
  search: ''
}

/**
 * Fetch tenders, returning a cached result when available. Pass `force` to
 * bypass and refresh the cache (e.g. the Refresh button).
 */
export function fetchTenders(q: TenderQuery, force = false): Promise<TenderResult> {
  const k = keyOf(q)
  if (force) cache.delete(k)
  let p = cache.get(k)
  if (!p) {
    p = api.searchTenders(q)
    // Drop failed lookups so a later attempt can retry rather than re-serving
    // the rejection forever.
    p.catch(() => cache.delete(k))
    cache.set(k, p)
  }
  return p
}

/** Warm the default Search Tender view in the background. Failures are ignored
 *  — the tab will surface any error itself when the user opens it. */
export function prefetchTenders(): void {
  void fetchTenders(DEFAULT_TENDER_QUERY).catch(() => {})
}

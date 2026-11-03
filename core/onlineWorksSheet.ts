import type { ExcelTable } from './types'

/**
 * One-way mirror of the Works List to a Google Sheet: each office gets its own tab, rewritten whole whenever that
 * office's Works List changes. Nothing is ever read back from the sheet.
 */

/** Sheets rejects tab titles longer than 100 characters, and \ / ? * [ ] : . */
export function officeTabTitle(officeKey: string): string {
  const parts = officeKey.split('|').map((p) => p.trim()).filter(Boolean)
  const title = (parts.length > 0 ? parts.join(' - ') : 'Works List').replace(/[\\/?*[\]:]/g, ' ').replace(/\s+/g, ' ').trim()
  return title.slice(0, 100) || 'Works List'
}

/** The grid written to the tab: the source sheet's own column order when known (as the Works List download does), else the table headers. */
export function worksTableToGrid(table: ExcelTable): string[][] {
  const headers = table.sourceHeaders && table.sourceHeaders.length > 0 ? table.sourceHeaders : table.headers
  return [headers, ...table.rows.map((r) => headers.map((h) => (r[h] ?? '').toString()))]
}

/** Cheap change signature (FNV-1a over the grid) so a save that didn't touch the Works List pushes nothing. */
export function gridSignature(grid: string[][]): string {
  let h = 0x811c9dc5
  for (const row of grid) {
    for (const cell of row) {
      for (let i = 0; i < cell.length; i++) {
        h ^= cell.charCodeAt(i)
        h = Math.imul(h, 0x01000193) >>> 0
      }
      h ^= 0x1f
      h = Math.imul(h, 0x01000193) >>> 0
    }
    h ^= 0x1e
    h = Math.imul(h, 0x01000193) >>> 0
  }
  return `${grid.length}:${h.toString(16)}`
}

/** Splits the rows into batches so each Sheets request stays well under the payload limit. */
export function chunkRows<T>(rows: T[], size: number): T[][] {
  const out: T[][] = []
  for (let i = 0; i < rows.length; i += size) out.push(rows.slice(i, i + size))
  return out
}

/** A1 column letters for a 1-based column index (1 → A, 27 → AA). */
export function columnLetters(n: number): string {
  let s = ''
  for (let c = Math.max(1, n); c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s
  return s
}

/** Quotes a tab title for use inside an A1 range. */
export function quoteTab(title: string): string {
  return `'${title.replace(/'/g, "''")}'`
}

/** The Works List table (the first table) of every office in the workspace, keyed by office. */
export function worksGridsByOffice(state: {
  tablesByOffice?: Record<string, ExcelTable[]>
  tables: ExcelTable[]
  currentOfficeKey?: string
}): Record<string, string[][]> {
  const out: Record<string, string[][]> = {}
  for (const [key, tables] of Object.entries(state.tablesByOffice ?? {})) {
    if (tables[0]) out[key] = worksTableToGrid(tables[0])
  }
  if (state.currentOfficeKey && state.tables[0]) out[state.currentOfficeKey] = worksTableToGrid(state.tables[0])
  return out
}

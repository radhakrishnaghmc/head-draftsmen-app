import type { ExcelTable } from '@core/types'
import type { PlaceholderMatch } from '@core/createDocument'

/**
 * The standard "works" schema for the tender/agreement workflow. These columns
 * come from the APP template Excel and define the built-in Work database — the
 * user fills rows in-app (no Excel upload needed).
 */
export const WORKS_COLUMNS: string[] = [
  'Zone',
  'Circle',
  'CNO',
  'Wincode',
  'Name of the work',
  'Amount of estimate',
  'Estimate Amount ECV',
  'Contract Amount',
  'Sanction By',
  'Tender notice',
  'Financial Year',
  'Tender Notice No',
  'Tender notice Date',
  'Tender ID',
  'Tender Percentage',
  'Name of the Agency',
  'Address of the agency',
  'Phone number of the agency',
  'TCV',
  'Intimation Date',
  'EMD 1.5%',
  'EMD 1%',
  'ASD',
  'Reservation',
  'Completion Period',
  'Agmt Date',
  'Technical Sanc No',
  'TS date'
]

/** Build a fresh, empty works database (one blank row to type into). */
export function createWorksTable(): ExcelTable {
  const blankRow: Record<string, string> = {}
  for (const h of WORKS_COLUMNS) blankRow[h] = ''
  return {
    id: `works-${Date.now()}`,
    name: 'Works database',
    path: '',
    headers: [...WORKS_COLUMNS],
    rows: [blankRow]
  }
}

/**
 * Force a table onto the standard works schema: guarantee every standard
 * column exists, while keeping any *extra* columns already on the table
 * (e.g. ones added by a previous import via applyWorksSchemaWithMapping) —
 * only ever adding to the schema, never silently dropping columns the user
 * already has. Guarantees at least one (blank) row.
 */
export function applyWorksSchema(table: ExcelTable): ExcelTable {
  const extraHeaders = table.headers.filter((h) => !WORKS_COLUMNS.includes(h))
  const headers = [...WORKS_COLUMNS, ...extraHeaders]
  const rows = (table.rows.length > 0 ? table.rows : [{}]).map((row) => {
    const next: Record<string, string> = {}
    for (const h of headers) next[h] = row[h] ?? ''
    return next
  })
  return { ...table, path: '', headers, rows }
}

/**
 * Like applyWorksSchema, but pulls each standard column's value from
 * whichever imported header `mapping` resolved it to (semantic/keyword
 * match — see core/createDocument.ts's matchPlaceholdersToColumns), not an
 * exact name match. A freshly-downloaded sheet using different column names
 * (e.g. "Estimate Amount" instead of "Amount of estimate") lands in the
 * right place instead of being silently dropped. Falls back to an exact
 * name match for any column the mapping didn't resolve.
 *
 * Any imported column the mapping didn't claim for a standard column (e.g.
 * "Eoffice", "Download start time" — real fields the app doesn't already
 * track) is kept too, appended after the standard columns under its own
 * original name, rather than discarded.
 */
export function applyWorksSchemaWithMapping(
  importedHeaders: string[],
  importedRows: Record<string, string>[],
  mapping: PlaceholderMatch[],
  meta: { id: string; name: string; path: string }
): ExcelTable {
  const columnFor = new Map(mapping.map((m) => [m.label, m.column]))
  const claimed = new Set(mapping.map((m) => m.column).filter((c): c is string => c !== null))
  const extraHeaders = importedHeaders.filter((h) => !claimed.has(h))
  const headers = [...WORKS_COLUMNS, ...extraHeaders]

  const rows = (importedRows.length > 0 ? importedRows : [{}]).map((row) => {
    const next: Record<string, string> = {}
    for (const h of WORKS_COLUMNS) {
      const matchedCol = columnFor.get(h)
      next[h] = (matchedCol ? row[matchedCol] : undefined) ?? row[h] ?? ''
    }
    for (const h of extraHeaders) {
      next[h] = row[h] ?? ''
    }
    return next
  })
  return { ...meta, headers, rows }
}

import type { ExcelTable } from './types'

/**
 * A raw sheet grid: the first sheet of a workbook read as an array-of-arrays of
 * strings, before any header row has been chosen. The renderer shows this in the
 * Header Row Picker so the user can confirm which row holds the column names.
 */
export interface SheetGrid {
  id: string
  name: string
  path: string
  sheetName: string
  grid: string[][]
  /** 0-based index of the first grid row within the real sheet (Excel row = startRow + i + 1). */
  startRow: number
  /** True when this sheet is hidden (or very hidden) in the workbook — the user never sees it in Excel. */
  hidden?: boolean
}

const cellStr = (v: unknown): string => (v === undefined || v === null ? '' : String(v))

function isNumericCell(s: string): boolean {
  const n = s.replace(/,/g, '')
  return Number.isFinite(Number(n))
}

/**
 * Best guess for the header row: the row with the most non-empty, non-numeric
 * cells within the first several rows. This skips a merged title row (which
 * has only one cell) and lands on the real column-header row. Cells that look
 * like plain numbers are excluded from the count — a data row that happens to
 * be fully populated (e.g. quantities and rates in every column) can otherwise
 * outscore the true header row above it, which is text-only and may have a
 * genuinely blank column (an unlabeled serial-number column, say). Returns a
 * 0-based index.
 */
export function guessHeaderRow(grid: string[][]): number {
  const limit = Math.min(grid.length, 15)
  let bestIdx = 0
  let bestCount = -1
  for (let r = 0; r < limit; r++) {
    const count = (grid[r] ?? []).filter((c) => {
      const s = cellStr(c).trim()
      return s.length > 0 && !isNumericCell(s)
    }).length
    if (count > bestCount) {
      bestCount = count
      bestIdx = r
    }
  }
  return bestIdx
}

/**
 * Build a flat table from a raw grid given the chosen header row index.
 * Blank header cells that still have data below them are named "Column N";
 * fully-empty columns are dropped. Duplicate header names get a numeric suffix.
 * Fully-empty data rows are skipped. Every cell is coerced to a string.
 */
export function buildTableFromGrid(
  grid: string[][],
  headerRowIndex: number,
  meta: { id: string; name: string; path: string }
): ExcelTable {
  if (grid.length === 0) {
    return { id: meta.id, name: meta.name, path: meta.path, headers: [], rows: [] }
  }
  const headerIdx = Math.max(0, Math.min(headerRowIndex, grid.length - 1))
  const headerRow = grid[headerIdx] ?? []
  const dataRows = grid.slice(headerIdx + 1)
  const width = grid.reduce((max, row) => Math.max(max, row.length), 0)

  const usedNames = new Set<string>()
  const kept: { index: number; name: string }[] = []
  for (let c = 0; c < width; c++) {
    const rawHeader = cellStr(headerRow[c]).trim()
    const hasData = dataRows.some((row) => cellStr(row[c]).trim().length > 0)
    if (rawHeader === '' && !hasData) continue // fully-empty column: drop
    let name = rawHeader !== '' ? rawHeader : `Column ${c + 1}`
    if (usedNames.has(name)) {
      let n = 2
      while (usedNames.has(`${name} (${n})`)) n += 1
      name = `${name} (${n})`
    }
    usedNames.add(name)
    kept.push({ index: c, name })
  }

  const headers = kept.map((k) => k.name)
  const rows: Record<string, string>[] = []
  for (const row of dataRows) {
    const rowObj: Record<string, string> = {}
    let hasValue = false
    for (const k of kept) {
      const value = cellStr(row[k.index])
      rowObj[k.name] = value
      if (value.trim().length > 0) hasValue = true
    }
    if (hasValue) rows.push(rowObj)
  }

  return { id: meta.id, name: meta.name, path: meta.path, headers, rows }
}

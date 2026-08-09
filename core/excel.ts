import * as fs from 'fs'
import * as path from 'path'
import * as XLSX from 'xlsx'
import type { ExcelTable } from './types'
import { buildTableFromGrid, guessHeaderRow, type SheetGrid } from './sheet'

let counter = 0
function nextId(prefix: string): string {
  counter += 1
  return `${prefix}_${Date.now().toString(36)}_${counter}`
}

const cellStr = (v: unknown): string => (v === undefined || v === null ? '' : String(v))

/**
 * Pick the first *visible* sheet. Workbooks often keep a hidden summary/config
 * sheet at index 0, so `SheetNames[0]` can point at a sheet the user never sees.
 * Falls back to the first sheet if visibility metadata is missing or all are hidden.
 */
function firstVisibleSheetName(workbook: XLSX.WorkBook): string | undefined {
  const meta = workbook.Workbook?.Sheets
  if (meta) {
    for (let i = 0; i < workbook.SheetNames.length; i++) {
      // Hidden: 0 = visible, 1 = hidden, 2 = very hidden.
      if (!meta[i]?.Hidden) return workbook.SheetNames[i]
    }
  }
  return workbook.SheetNames[0]
}

// Read cells directly from the decoded range so every row (including blank
// ones) keeps its real position — the picker can then show Excel's own row
// numbers instead of a compacted sequence.
function gridFromSheet(sheet: XLSX.WorkSheet): { grid: string[][]; startRow: number } {
  const ref = sheet['!ref']
  if (!ref) return { grid: [], startRow: 0 }
  const range = XLSX.utils.decode_range(ref)
  const grid: string[][] = []
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: string[] = []
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })]
      row.push(cell ? cellStr(cell.w ?? cell.v) : '')
    }
    grid.push(row)
  }
  return { grid, startRow: range.s.r }
}

/**
 * Read the first visible sheet of a workbook into a raw grid of strings (no
 * header chosen yet). Used by the Header Row Picker so the user can confirm
 * which row holds the column names.
 */
export function readExcelGrid(filePath: string): SheetGrid {
  const buffer = fs.readFileSync(filePath)
  const workbook = XLSX.read(buffer, { type: 'buffer' })
  const sheetName = firstVisibleSheetName(workbook)
  const base = {
    id: nextId('xls'),
    name: path.basename(filePath),
    path: filePath,
    sheetName: sheetName ?? ''
  }
  if (!sheetName) return { ...base, grid: [], startRow: 0 }
  return { ...base, ...gridFromSheet(workbook.Sheets[sheetName]) }
}

/**
 * Read every sheet of a workbook buffer into its own raw grid — shared by
 * the local-file and downloaded-link paths (see readAllSheetGrids and
 * core/googleImport.ts's importAllSheetsFromGoogleLink).
 */
export function allSheetGridsFromBuffer(buffer: Buffer, name: string, filePath: string): SheetGrid[] {
  const workbook = XLSX.read(buffer, { type: 'buffer' })
  // Per-sheet visibility metadata (Hidden: 0 = visible, 1 = hidden, 2 = very
  // hidden), indexed by sheet position — so callers can flag hidden sheets.
  const meta = workbook.Workbook?.Sheets
  return workbook.SheetNames.map((sheetName, i) => {
    const { grid, startRow } = gridFromSheet(workbook.Sheets[sheetName])
    return { id: nextId('xls'), name, path: filePath, sheetName, grid, startRow, hidden: !!meta?.[i]?.Hidden }
  })
}

/**
 * Read every sheet of a workbook into its own raw grid — used for reference
 * workbooks (e.g. a rates database) that spread their data across many
 * sheets, unlike the single-sheet uploads the rest of the app deals with.
 */
export function readAllSheetGrids(filePath: string): SheetGrid[] {
  const buffer = fs.readFileSync(filePath)
  return allSheetGridsFromBuffer(buffer, path.basename(filePath), filePath)
}

/**
 * Parse an Excel workbook (.xlsx/.xls) into a flat table with the header row
 * auto-detected. Retained for tests and any non-interactive callers.
 */
export function parseExcelBuffer(buffer: Buffer, fileName: string, filePath = ''): ExcelTable {
  const workbook = XLSX.read(buffer, { type: 'buffer' })
  const sheetName = firstVisibleSheetName(workbook)
  const meta = { id: nextId('xls'), name: fileName, path: filePath }
  if (!sheetName) return { ...meta, headers: [], rows: [] }
  const sheet = workbook.Sheets[sheetName]
  const aoa = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    defval: '',
    blankrows: false,
    raw: false
  })
  const grid = aoa.map((row) => row.map(cellStr))
  return buildTableFromGrid(grid, guessHeaderRow(grid), meta)
}

/** Read an Excel file from disk and parse it (header auto-detected). */
export function parseExcelFile(filePath: string): ExcelTable {
  const buffer = fs.readFileSync(filePath)
  return parseExcelBuffer(buffer, path.basename(filePath), filePath)
}

/** A small preview of one worksheet — its name plus the top-left corner of its
 * cells — used to draw the Excel Sheet Separator's live sheet tiles. */
export interface SheetPreview {
  name: string
  /** Whether the sheet is hidden in the workbook (shown faded, still separable). */
  hidden: boolean
  /** Up to `maxRows` × `maxCols` of stringified cell values from the top-left. */
  rows: string[][]
}

/**
 * Read a lightweight preview of every sheet in a workbook (name + a small
 * top-left block of cells), for the sheet tiles. `sheetRows` caps how much of
 * each sheet is parsed so even a large workbook previews quickly.
 */
export function readSheetPreviews(filePath: string, maxRows = 7, maxCols = 5): SheetPreview[] {
  const buffer = fs.readFileSync(filePath)
  const workbook = XLSX.read(buffer, { type: 'buffer', sheetRows: maxRows, cellFormula: false })
  const meta = workbook.Workbook?.Sheets
  return workbook.SheetNames.map((name, i) => {
    const ws = workbook.Sheets[name]
    const aoa = ws
      ? (XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: '' }) as unknown[][])
      : []
    const rows = aoa.slice(0, maxRows).map((r) => r.slice(0, maxCols).map(cellStr))
    return { name, hidden: !!meta?.[i]?.Hidden, rows }
  })
}

/** Serialize a table back to an .xlsx workbook buffer, e.g. for saving to disk. */
export function buildWorkbookBuffer(table: ExcelTable): Buffer {
  const aoa = [table.headers, ...table.rows.map((row) => table.headers.map((h) => row[h] ?? ''))]
  const sheet = XLSX.utils.aoa_to_sheet(aoa)
  const workbook = XLSX.utils.book_new()
  // Excel sheet names cap at 31 chars.
  XLSX.utils.book_append_sheet(workbook, sheet, (table.name || 'Sheet1').slice(0, 31))
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer
}

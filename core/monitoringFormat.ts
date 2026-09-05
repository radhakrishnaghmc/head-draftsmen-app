import type { SheetGrid } from './sheet'

/** The subset of the app's Office shape (src/office.ts) this module needs — kept local so core/ stays framework-agnostic. */
export interface OfficeLike {
  zone?: string
  circle?: string
  circleNumber?: string
}

/** One No./Amt pair, as every column group in the Monitoring Format sheet is laid out. */
export interface MonitoringFormatBucket {
  no: number
  amt: number
}

/** One row of the Monitoring Format table — either a single item type (CC Roads, SWD, …) or the sheet's own "Total" row. */
export interface MonitoringFormatRow {
  itemType: string
  totalWorks: MonitoringFormatBucket
  completed: MonitoringFormatBucket
  upto25: MonitoringFormatBucket
  upto50: MonitoringFormatBucket
  upto75: MonitoringFormatBucket
  above75: MonitoringFormatBucket
  progressTotal: MonitoringFormatBucket
  toBeStarted: MonitoringFormatBucket
  tenderProcess: MonitoringFormatBucket
  heldUp: MonitoringFormatBucket
  cancelled: MonitoringFormatBucket
}

export interface MonitoringFormatSummary {
  /** The office label read off the sheet itself (e.g. "58-Nizampet" or "Quthbullapur Zone"), not the app's own office picker. */
  officeLabel: string
  sheetName: string
  /** The sheet's own "Date:-" value, if present. */
  asOfDate?: string
  /** One row per item type (CC Roads, BT Roads, SWD, …), in sheet order. */
  rows: MonitoringFormatRow[]
  /** The sheet's own bottom "Total" row — every tile on the Dashboard is built from this. */
  totals: MonitoringFormatRow
}

const ZERO_BUCKET: MonitoringFormatBucket = { no: 0, amt: 0 }

// Fixed column layout of the Monitoring Format template (0-indexed): every
// group is a No./Amt pair, in this order. Confirmed against a real workbook
// (both the zone-wide "QBZ MF" sheet and a per-circle "C58 MF" sheet use the
// identical layout), so it's read positionally rather than by re-matching
// header text on every row.
const COLS = {
  itemType: 0,
  totalWorks: 1,
  completed: 3,
  upto25: 5,
  upto50: 7,
  upto75: 9,
  above75: 11,
  progressTotal: 13,
  toBeStarted: 15,
  tenderProcess: 17,
  heldUp: 19,
  cancelled: 21
} as const

function toNumber(cell: string | undefined): number {
  if (!cell) return 0
  const n = Number(String(cell).replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : 0
}

function bucketAt(row: string[], startCol: number): MonitoringFormatBucket {
  return { no: toNumber(row[startCol]), amt: toNumber(row[startCol + 1]) }
}

function addBuckets(a: MonitoringFormatBucket, b: MonitoringFormatBucket): MonitoringFormatBucket {
  return { no: a.no + b.no, amt: a.amt + b.amt }
}

function emptyRow(itemType: string): MonitoringFormatRow {
  return {
    itemType,
    totalWorks: { ...ZERO_BUCKET },
    completed: { ...ZERO_BUCKET },
    upto25: { ...ZERO_BUCKET },
    upto50: { ...ZERO_BUCKET },
    upto75: { ...ZERO_BUCKET },
    above75: { ...ZERO_BUCKET },
    progressTotal: { ...ZERO_BUCKET },
    toBeStarted: { ...ZERO_BUCKET },
    tenderProcess: { ...ZERO_BUCKET },
    heldUp: { ...ZERO_BUCKET },
    cancelled: { ...ZERO_BUCKET }
  }
}

function sumRows(rows: MonitoringFormatRow[]): MonitoringFormatRow {
  return rows.reduce((sum, r) => ({
    itemType: 'Total',
    totalWorks: addBuckets(sum.totalWorks, r.totalWorks),
    completed: addBuckets(sum.completed, r.completed),
    upto25: addBuckets(sum.upto25, r.upto25),
    upto50: addBuckets(sum.upto50, r.upto50),
    upto75: addBuckets(sum.upto75, r.upto75),
    above75: addBuckets(sum.above75, r.above75),
    progressTotal: addBuckets(sum.progressTotal, r.progressTotal),
    toBeStarted: addBuckets(sum.toBeStarted, r.toBeStarted),
    tenderProcess: addBuckets(sum.tenderProcess, r.tenderProcess),
    heldUp: addBuckets(sum.heldUp, r.heldUp),
    cancelled: addBuckets(sum.cancelled, r.cancelled)
  }), emptyRow('Total'))
}

/** Read the "Name of the Cir(c)le:-" / "Date:-" label+value pairs off the sheet's first few rows. */
function readLabelledValues(grid: string[][]): { office?: string; date?: string } {
  let office: string | undefined
  let date: string | undefined
  for (const row of grid.slice(0, 6)) {
    for (let i = 0; i < row.length; i++) {
      const cell = (row[i] || '').trim().toLowerCase()
      if (!cell) continue
      const next = row.slice(i + 1).find((v) => (v || '').trim() !== '')
      if (!office && cell.startsWith('name of the cir')) office = next?.trim()
      else if (!date && cell.startsWith('date')) date = next?.trim()
    }
  }
  return { office, date }
}

/** Parse one "MF" sheet's item-type table into per-item-type rows plus the sheet's own Total row. */
export function parseMonitoringFormatSheet(sheet: SheetGrid): MonitoringFormatSummary {
  const grid = sheet.grid
  const headerRowIdx = grid.findIndex((row) => (row[COLS.itemType] || '').trim().toLowerCase() === 'item type')
  if (headerRowIdx === -1) {
    throw new Error(`"${sheet.sheetName}" doesn't look like a Monitoring Format sheet (no "Item type" header row found).`)
  }

  const { office, date } = readLabelledValues(grid)

  const rows: MonitoringFormatRow[] = []
  let totalsRow: MonitoringFormatRow | undefined
  // Data starts two rows below the "Item type" header (it spans 3 header rows: group / sub-group / No.-Amt.).
  for (let i = headerRowIdx + 3; i < grid.length; i++) {
    const row = grid[i]
    const itemType = (row[COLS.itemType] || '').trim()
    if (!itemType) break
    const parsed: MonitoringFormatRow = {
      itemType,
      totalWorks: bucketAt(row, COLS.totalWorks),
      completed: bucketAt(row, COLS.completed),
      upto25: bucketAt(row, COLS.upto25),
      upto50: bucketAt(row, COLS.upto50),
      upto75: bucketAt(row, COLS.upto75),
      above75: bucketAt(row, COLS.above75),
      progressTotal: bucketAt(row, COLS.progressTotal),
      toBeStarted: bucketAt(row, COLS.toBeStarted),
      tenderProcess: bucketAt(row, COLS.tenderProcess),
      heldUp: bucketAt(row, COLS.heldUp),
      cancelled: bucketAt(row, COLS.cancelled)
    }
    if (itemType.toLowerCase() === 'total') {
      totalsRow = parsed
      break
    }
    rows.push(parsed)
  }

  return {
    officeLabel: office || sheet.sheetName,
    sheetName: sheet.sheetName,
    asOfDate: date,
    rows,
    // Most Monitoring Format workbooks already carry their own "Total" row —
    // fall back to summing the item-type rows for the rare one that doesn't.
    totals: totalsRow || sumRows(rows)
  }
}

/**
 * Pick the one sheet in a Monitoring Format workbook that belongs to the
 * current office, out of every circle's (and the zone's) sheet. Matched by
 * the sheet's own "Name of the Cir(c)le:-" cell rather than the sheet's tab
 * name, since tab names/abbreviations aren't standardised — the label cell
 * always spells out either "<CNO>-<Circle>" or "<Zone> Zone".
 */
export function findMonitoringFormatSheet(sheets: SheetGrid[], office: OfficeLike): SheetGrid | null {
  const candidates = sheets.filter(
    (s) => /\bmf\b/i.test(s.sheetName) && !/list/i.test(s.sheetName)
  )
  if (candidates.length === 0) return null

  const withLabel = candidates.map((s) => ({ sheet: s, label: (readLabelledValues(s.grid).office || '').toLowerCase() }))

  if (office.circle) {
    const cno = (office.circleNumber || '').trim()
    const circle = office.circle.trim().toLowerCase()
    const byLabel = withLabel.find(
      ({ label }) => label && ((cno && label.startsWith(`${cno}-`)) || label.includes(circle))
    )
    if (byLabel) return byLabel.sheet
    // Fall back to the tab name itself carrying the circle number, e.g. "C58 MF".
    if (cno) {
      const byName = candidates.find((s) => new RegExp(`(^|[^0-9])0*${cno}([^0-9]|$)`).test(s.sheetName))
      if (byName) return byName
    }
    return null
  }

  if (office.zone) {
    const zone = office.zone.trim().toLowerCase()
    const byLabel = withLabel.find(({ label }) => label && label.includes(zone) && !/^\d+-/.test(label))
    if (byLabel) return byLabel.sheet
    // A zone-only office has no circle number. Its tab is often just an
    // abbreviation (e.g. "QBZ MF") that won't textually match the zone name,
    // so among the "MF" sheets that don't look like a per-circle sheet
    // ("C##"), fall back to the first one — the zone rollup is conventionally
    // the first sheet in these workbooks.
    const byName = candidates.filter((s) => !/^c\d/i.test(s.sheetName.trim()))
    if (byName.length > 0) return byName[0]
  }

  return null
}

/** Download-and-parse convenience: pick this office's sheet out of the workbook and parse it. */
export function extractMonitoringFormatForOffice(sheets: SheetGrid[], office: OfficeLike): MonitoringFormatSummary {
  const sheet = findMonitoringFormatSheet(sheets, office)
  if (!sheet) {
    throw new Error(
      "Couldn't find this office's sheet in that Monitoring Format workbook — check it has a sheet for " +
        (office.circle ? `circle "${office.circle}"` : `zone "${office.zone}"`) +
        ' with a "Name of the Cir(c)le:-" label matching it.'
    )
  }
  return parseMonitoringFormatSheet(sheet)
}

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

/** The six per-status buckets shown as clickable tiles on the Dashboard ("totalWorks" and the progress sub-buckets aren't statuses of their own). */
export type MfStatusKey = 'completed' | 'progressTotal' | 'toBeStarted' | 'tenderProcess' | 'heldUp' | 'cancelled'

/** One row of a circle's "list of works" sheet — the detail behind the Abstract's per-status counts. */
export interface MonitoringFormatWorkRow {
  slNo: string
  circle: string
  ward: string
  workName: string
  estimateAmt: number
  typeOfWork: string
  sourceOfSanction: string
  sanctionDate: string
  agreementDate: string
  targetDate: string
  agencyDetails: string
  /** Raw "Current status" cell text, as typed in the sheet (casing/wording varies by circle). */
  status: string
  /** `status`, normalized to whichever Dashboard tile it belongs under. */
  statusKey: MfStatusKey
  winCode: string
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
  /**
   * The same Abstract sheet's second pivot, one row per ward instead of per
   * item type (each row's `itemType` field holds the ward's name/label) —
   * undefined when the sheet doesn't have one, so callers can tell "no ward
   * block on this sheet" apart from "block present but happens to be empty".
   */
  wardRows?: MonitoringFormatRow[]
  /** The ward-wise block's own bottom "Total" row, mirroring `totals` for item type. */
  wardTotals?: MonitoringFormatRow
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

/**
 * Parses one pivot block sharing the Abstract sheet's fixed column layout
 * (COLS): a header row whose first cell reads `headerLabel` (e.g. "Item
 * type" or "Ward"), 3 header rows deep, then one data row per group ending
 * in the block's own "Total" row. Searches from `searchFrom` onward so a
 * second block (the ward-wise pivot) is found below the first (the
 * item-type one) rather than re-matching it. Returns null when no row with
 * that header label exists from `searchFrom` on — the sheet has no such block.
 */
function parseAbstractBlock(
  grid: string[][],
  searchFrom: number,
  headerLabel: string
): { rows: MonitoringFormatRow[]; totals: MonitoringFormatRow; endIdx: number } | null {
  let headerRowIdx = -1
  for (let i = searchFrom; i < grid.length; i++) {
    if ((grid[i][COLS.itemType] || '').trim().toLowerCase() === headerLabel) {
      headerRowIdx = i
      break
    }
  }
  if (headerRowIdx === -1) return null

  const rows: MonitoringFormatRow[] = []
  let totalsRow: MonitoringFormatRow | undefined
  let endIdx = grid.length
  // Data starts two rows below the header (it spans 3 header rows: group / sub-group / No.-Amt.).
  for (let i = headerRowIdx + 3; i < grid.length; i++) {
    const row = grid[i]
    const label = (row[COLS.itemType] || '').trim()
    if (!label) {
      endIdx = i
      break
    }
    const parsed: MonitoringFormatRow = {
      itemType: label,
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
    if (label.toLowerCase() === 'total') {
      totalsRow = parsed
      endIdx = i + 1
      break
    }
    rows.push(parsed)
  }

  // Most Monitoring Format workbooks already carry their own "Total" row —
  // fall back to summing the group rows for the rare block that doesn't.
  return { rows, totals: totalsRow || sumRows(rows), endIdx }
}

/** Parse one "MF" sheet's item-type table (and, if present, its ward-wise pivot) into per-item-type/per-ward rows plus each block's own Total row. */
export function parseMonitoringFormatSheet(sheet: SheetGrid): MonitoringFormatSummary {
  const grid = sheet.grid
  const itemBlock = parseAbstractBlock(grid, 0, 'item type')
  if (!itemBlock) {
    throw new Error(`"${sheet.sheetName}" doesn't look like a Monitoring Format sheet (no "Item type" header row found).`)
  }
  const wardBlock = parseAbstractBlock(grid, itemBlock.endIdx, 'ward')

  const { office, date } = readLabelledValues(grid)

  return {
    officeLabel: office || sheet.sheetName,
    sheetName: sheet.sheetName,
    asOfDate: date,
    rows: itemBlock.rows,
    totals: itemBlock.totals,
    wardRows: wardBlock?.rows,
    wardTotals: wardBlock?.totals
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

/**
 * Free-text "Current status" cells vary by circle ("Held up" vs "held up",
 * one circle even uses "agreement to be concluded"), so match by keyword
 * rather than exact value. "agreement to be concluded" is sanctioned work
 * whose agreement hasn't been signed yet, i.e. not effectively started —
 * closest Dashboard bucket is "To Be Started".
 */
function normalizeWorkStatus(raw: string): MfStatusKey | null {
  const s = raw.trim().toLowerCase()
  if (!s) return null
  if (s.includes('cancel')) return 'cancelled'
  if (s.includes('held up') || s.includes('hold')) return 'heldUp'
  if (s.includes('tender')) return 'tenderProcess'
  if (s.includes('complete')) return 'completed'
  if (s.includes('progress')) return 'progressTotal'
  if (s.includes('to be started') || s.includes('not started') || s.includes('yet to') || s.includes('agreement')) {
    return 'toBeStarted'
  }
  return null
}

function colIndex(header: string[], test: (h: string) => boolean): number {
  return header.findIndex((h) => test((h || '').trim().toLowerCase()))
}

function cellAt(row: string[], idx: number): string {
  return idx >= 0 ? (row[idx] || '').trim() : ''
}

/**
 * Parse one circle's "list of works" sheet into individual work rows, each
 * tagged with the normalized status bucket it belongs to on the Dashboard.
 * Columns are located by header text (not fixed positions) since trailing
 * columns after "Current status" vary between circles.
 */
export function parseMonitoringFormatWorkSheet(sheet: SheetGrid): MonitoringFormatWorkRow[] {
  const grid = sheet.grid
  const headerRowIdx = grid.findIndex((row) => row.some((c) => (c || '').trim().toLowerCase() === 'current status'))
  if (headerRowIdx === -1) return []
  const header = grid[headerRowIdx]

  const col = {
    slNo: colIndex(header, (h) => h.startsWith('sl.no') || h.startsWith('sl no')),
    circle: colIndex(header, (h) => h.includes('circle')),
    ward: colIndex(header, (h) => h.includes('ward')),
    workName: colIndex(header, (h) => /name.*work/.test(h)),
    estimateAmt: colIndex(header, (h) => h.includes('estimate')),
    typeOfWork: colIndex(header, (h) => h.includes('type') && h.includes('work')),
    sourceOfSanction: colIndex(header, (h) => h.includes('source')),
    sanctionDate: colIndex(header, (h) => h.includes('sanction') && h.includes('date')),
    agreementDate: colIndex(header, (h) => h.includes('agreement')),
    targetDate: colIndex(header, (h) => h.includes('target')),
    agencyDetails: colIndex(header, (h) => h.includes('agency')),
    status: colIndex(header, (h) => h.includes('current status')),
    winCode: colIndex(header, (h) => h.includes('win'))
  }

  const rows: MonitoringFormatWorkRow[] = []
  for (let i = headerRowIdx + 1; i < grid.length; i++) {
    const row = grid[i]
    const slNo = cellAt(row, col.slNo)
    const workName = cellAt(row, col.workName)
    if (!slNo && !workName) continue
    const status = cellAt(row, col.status)
    const statusKey = normalizeWorkStatus(status)
    if (!statusKey) continue
    rows.push({
      slNo,
      circle: cellAt(row, col.circle),
      ward: cellAt(row, col.ward),
      workName,
      estimateAmt: toNumber(row[col.estimateAmt]),
      typeOfWork: cellAt(row, col.typeOfWork),
      sourceOfSanction: cellAt(row, col.sourceOfSanction),
      sanctionDate: cellAt(row, col.sanctionDate),
      agreementDate: cellAt(row, col.agreementDate),
      targetDate: cellAt(row, col.targetDate),
      agencyDetails: cellAt(row, col.agencyDetails),
      status,
      statusKey,
      winCode: cellAt(row, col.winCode)
    })
  }
  return rows
}

/**
 * Pick the "list of works" sheet(s) belonging to this office. A circle office
 * has exactly one (matched the same way as the MF sheet's tab-name fallback:
 * by circle number embedded in the tab name, e.g. "C58 list of works"). A
 * zone-only office has no single sheet to pick — there's no zone-wide "list
 * of works" rollup, only the zone-wide MF abstract — so every circle's list
 * in the workbook is in scope (the whole workbook already belongs to one zone).
 */
export function findMonitoringFormatWorkSheets(sheets: SheetGrid[], office: OfficeLike): SheetGrid[] {
  const candidates = sheets.filter((s) => /list/i.test(s.sheetName))
  if (candidates.length === 0) return []

  if (office.circle) {
    const cno = (office.circleNumber || '').trim()
    const circle = office.circle.trim().toLowerCase()
    const match = candidates.find((s) => {
      if (cno && new RegExp(`(^|[^0-9])0*${cno}([^0-9]|$)`).test(s.sheetName)) return true
      return circle ? s.sheetName.toLowerCase().includes(circle) : false
    })
    return match ? [match] : []
  }

  return candidates
}

/** Download-and-parse convenience for the per-status work lists behind the Dashboard's status tiles. */
export function extractMonitoringFormatWorksForOffice(sheets: SheetGrid[], office: OfficeLike): MonitoringFormatWorkRow[] {
  return findMonitoringFormatWorkSheets(sheets, office).flatMap(parseMonitoringFormatWorkSheet)
}

/** Works matching one Dashboard status tile (Completed / In Progress / To Be Started / …). */
export function filterMonitoringFormatWorksByStatus(
  works: MonitoringFormatWorkRow[],
  key: MfStatusKey
): MonitoringFormatWorkRow[] {
  return works.filter((w) => w.statusKey === key)
}

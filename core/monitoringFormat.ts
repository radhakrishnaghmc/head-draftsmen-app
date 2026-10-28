import type { SheetGrid } from './sheet'

/** The subset of the app's Office shape (src/office.ts) this module needs — kept local so core/ stays framework-agnostic. */
export interface OfficeLike {
  zone?: string
  circle?: string
  circleNumber?: string
  /** For a zone-level office: the circle numbers (CNO) under that zone — lets a workbook that holds EVERY zone's data pick out just this zone's works. */
  circleNumbers?: string[]
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
  /** Every non-blank-header column of the sheet row, in the sheet's own order (parallel to `headers`). Absent on rows saved before this was added. */
  cells?: string[]
  /** The sheet's column headings, same order as `cells` — stored on the first row of each sheet only, to keep the synced payload small. */
  headers?: string[]
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
   * The same Abstract sheet's second pivot, one row per ward (a circle-level
   * sheet) or per circle (a zone-level sheet) instead of per item type (each
   * row's `itemType` field holds the group's name/label) — undefined when
   * the sheet doesn't have one, so callers can tell "no second block on this
   * sheet" apart from "block present but happens to be empty".
   */
  wardRows?: MonitoringFormatRow[]
  /** The second block's own bottom "Total" row, mirroring `totals` for item type. */
  wardTotals?: MonitoringFormatRow
  /** Which grouping `wardRows`/`wardTotals` actually holds, read off the block's own header cell. Undefined when there's no second block. */
  wardPivotKind?: 'ward' | 'circle'
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

function excelSerialToDmy(n: number): string {
  const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000)
  const p = (x: number) => String(x).padStart(2, '0')
  return `${p(d.getUTCDate())}-${p(d.getUTCMonth() + 1)}-${d.getUTCFullYear()}`
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
      else if (!date && (cell.startsWith('date') || cell === 'dt' || cell === 'dt:')) {
        const v = next?.trim()
        // A bare Excel serial (e.g. 46320) -> dd-mm-yyyy.
        date = v && /^\d{5}$/.test(v) ? excelSerialToDmy(Number(v)) : v
      }
    }
  }
  return { office, date }
}

// 2026-27 template ("Sl. No | Circle | Category | Sub | Status ..."): Category/Sub-category
// labels, Spillover + Current-year pairs ahead of the Total works / Completed / Progress /
// To be started / Held up / Cancelled / Tender process pairs. No progress-% buckets.
const NEW_COLS = {
  circle: 1,
  category: 2,
  sub: 3,
  totalWorks: 8,
  completed: 10,
  progressTotal: 12,
  toBeStarted: 14,
  heldUp: 16,
  cancelled: 18,
  tenderProcess: 20
} as const

// The same template without the Circle column (2025-26 zone/circle tabs of a multi-zone workbook): every column sits one to the left.
const NEW_COLS_NO_CIRCLE = {
  circle: -1,
  category: 1,
  sub: 2,
  totalWorks: 7,
  completed: 9,
  progressTotal: 11,
  toBeStarted: 13,
  heldUp: 15,
  cancelled: 17,
  tenderProcess: 19
} as const

const lc = (v: string | undefined) => (v || '').trim().toLowerCase()

function newLayoutHeaderIdx(grid: string[][]): number {
  return grid
    .slice(0, 10)
    .findIndex(
      (r) =>
        (lc(r[NEW_COLS.category]) === 'category' && lc(r[4]) === 'status') ||
        (lc(r[NEW_COLS_NO_CIRCLE.category]) === 'category' && lc(r[3]) === 'status')
    )
}

/** Which column set a new-layout sheet uses, from its header row. */
function newLayoutCols(grid: string[][], headerIdx: number): typeof NEW_COLS | typeof NEW_COLS_NO_CIRCLE {
  return lc(grid[headerIdx]?.[NEW_COLS.category]) === 'category' && lc(grid[headerIdx]?.[4]) === 'status' ? NEW_COLS : NEW_COLS_NO_CIRCLE
}

function newLayoutRow(label: string, row: string[], cols: { [K in 'totalWorks' | 'completed' | 'progressTotal' | 'toBeStarted' | 'heldUp' | 'cancelled' | 'tenderProcess']: number }): MonitoringFormatRow {
  return {
    ...emptyRow(label),
    totalWorks: bucketAt(row, cols.totalWorks),
    completed: bucketAt(row, cols.completed),
    progressTotal: bucketAt(row, cols.progressTotal),
    toBeStarted: bucketAt(row, cols.toBeStarted),
    heldUp: bucketAt(row, cols.heldUp),
    cancelled: bucketAt(row, cols.cancelled),
    tenderProcess: bucketAt(row, cols.tenderProcess)
  }
}

function parseNewLayoutSheet(sheet: SheetGrid, headerIdx: number): MonitoringFormatSummary {
  const grid = sheet.grid
  const COLS = newLayoutCols(grid, headerIdx)
  const rows: MonitoringFormatRow[] = []
  let totals: MonitoringFormatRow | undefined
  let category = ''
  let i = headerIdx + 4 // group row, No./Amount row, column-number row
  let firstCircle = ''
  for (; i < grid.length; i++) {
    const r = grid[i]
    const cat = (r[COLS.category] || '').trim()
    const sub = (r[COLS.sub] || '').trim()
    if (!firstCircle) firstCircle = (r[COLS.circle] || '').trim()
    if (lc(cat) === 'total') {
      totals = newLayoutRow('Total', r, COLS)
      i++
      break
    }
    if (!cat && !sub) break
    if (cat) category = cat
    rows.push(newLayoutRow(sub ? `${category} – ${sub}` : category, r, COLS))
  }
  totals = totals || sumRows(rows)

  // Second pivot. A circle sheet has a "Ward Name" block with the same column layout; the
  // zone sheet instead has per-category circle pivots (Roads / SWD / Others ...) whose rows are
  // "<circle no> | Total works ..." starting one column further left, which are summed per circle.
  let wardRows: MonitoringFormatRow[] | undefined
  let wardTotals: MonitoringFormatRow | undefined
  let wardPivotKind: 'ward' | 'circle' | undefined
  const wardHdr = grid.findIndex((r, k) => k >= i && lc(r[COLS.sub]) === 'ward name')
  if (wardHdr >= 0) {
    wardRows = []
    for (let k = wardHdr + 2; k < grid.length; k++) {
      const label = (grid[k][COLS.sub] || '').trim()
      if (lc(label) === 'total') {
        wardTotals = newLayoutRow('Total', grid[k], COLS)
        break
      }
      // Blank-named ward rows are an unfilled template (all zeros) — skip them.
      if (label) wardRows.push(newLayoutRow(label, grid[k], COLS))
      else if (k > wardHdr + 2 && !grid[k].some((c) => (c || '').trim())) break
    }
    if (wardRows.length === 0) wardRows = undefined
    else {
      wardTotals = wardTotals || sumRows(wardRows)
      wardPivotKind = 'ward'
    }
  } else {
    const PIVOT = COLS === NEW_COLS ? { totalWorks: 8, completed: 10, progressTotal: 12, toBeStarted: 14, heldUp: 16, cancelled: 18, tenderProcess: 20 } : { totalWorks: 7, completed: 9, progressTotal: 11, toBeStarted: 13, heldUp: 15, cancelled: 17, tenderProcess: 19 }
    const byCircle = new Map<string, MonitoringFormatRow>()
    let inPivot = false
    for (let k = i; k < grid.length; k++) {
      const label = (grid[k][COLS.sub + 4] || '').trim()
      if (/^(roads|swd|others)$/i.test(label)) inPivot = true
      else if (label && !/^(\d{1,4}|circle|total)$/i.test(label)) inPivot = false // BT / CC sub-pivots repeat Roads' rows
      else if (inPivot && /^\d{1,4}$/.test(label)) {
        const row = newLayoutRow(label, grid[k], PIVOT)
        const prev = byCircle.get(label)
        byCircle.set(label, prev ? sumRows([prev, row]) : row)
      }
    }
    if (byCircle.size > 0) {
      wardRows = [...byCircle.entries()]
        .sort((a, b) => Number(a[0]) - Number(b[0]))
        .map(([label, r]) => ({ ...r, itemType: label }))
      wardTotals = sumRows(wardRows)
      wardPivotKind = 'circle'
    }
  }

  const { office, date } = readLabelledValues(grid)
  return {
    officeLabel: office || firstCircle || sheet.sheetName,
    sheetName: sheet.sheetName,
    asOfDate: date,
    rows,
    totals,
    wardRows,
    wardTotals,
    wardPivotKind
  }
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
  // startsWith, not ===: real workbooks label the ward block "Ward Name and
  // number" rather than a bare "Ward", so an exact match silently missed it
  // and the ward-wise pivot never got parsed.
  let headerRowIdx = -1
  for (let i = searchFrom; i < grid.length; i++) {
    if ((grid[i][COLS.itemType] || '').trim().toLowerCase().startsWith(headerLabel)) {
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
  const newIdx = newLayoutHeaderIdx(grid)
  if (newIdx >= 0) return parseNewLayoutSheet(sheet, newIdx)
  const itemBlock = parseAbstractBlock(grid, 0, 'item type')
  if (!itemBlock) {
    throw new Error(`"${sheet.sheetName}" doesn't look like a Monitoring Format sheet (no "Item type" header row found).`)
  }
  // A zone-level sheet's second pivot breaks down by circle, not ward (a
  // zone has no wards of its own) — try both header labels rather than
  // assuming which one the sheet has.
  const wardBlock = parseAbstractBlock(grid, itemBlock.endIdx, 'ward')
  const circleBlock = wardBlock ? undefined : parseAbstractBlock(grid, itemBlock.endIdx, 'circle')
  const secondBlock = wardBlock || circleBlock

  const { office, date } = readLabelledValues(grid)

  return {
    officeLabel: office || sheet.sheetName,
    sheetName: sheet.sheetName,
    asOfDate: date,
    rows: itemBlock.rows,
    totals: itemBlock.totals,
    wardRows: secondBlock?.rows,
    wardTotals: secondBlock?.totals,
    wardPivotKind: wardBlock ? 'ward' : circleBlock ? 'circle' : undefined
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
  // The 2026-27 template's zone tab has no "MF" in its name ("LB Ngar Zone"), so a tab also
  // qualifies by its Category/Status header row.
  const candidates = sheets.filter(
    (s) => !/list/i.test(s.sheetName) && (/\bmf\b/i.test(s.sheetName) || newLayoutHeaderIdx(s.grid) >= 0)
  )
  if (candidates.length === 0) return null

  const withLabel = candidates.map((s) => {
    const { office } = readLabelledValues(s.grid)
    const idx = newLayoutHeaderIdx(s.grid)
    // New template: the circle (or zone) is named in the first data row's Circle cell ("11 - Nagole").
    const fromRow = idx >= 0 ? (s.grid[idx + 4]?.[newLayoutCols(s.grid, idx).circle] || '').trim() : ''
    return { sheet: s, label: (office || fromRow).toLowerCase() }
  })

  if (office.circle) {
    const cno = (office.circleNumber || '').trim()
    const circle = office.circle.trim().toLowerCase()
    // A workbook has several tabs per circle (MF, Spill AB, Current AB …): the "MF" one is the summary.
    const preferMf = (list: SheetGrid[]): SheetGrid | undefined => list.find((x) => /\bmf\b/i.test(x.sheetName)) ?? list[0]
    const byLabel = preferMf(
      withLabel
        .filter(({ label }) => label && ((cno && new RegExp(`^0*${cno}\\s*[-–]`).test(label)) || label.includes(circle)))
        .map(({ sheet }) => sheet)
    )
    if (byLabel) return byLabel
    // Fall back to the tab name itself carrying the circle number, e.g. "C58 MF".
    if (cno) {
      const byName = preferMf(candidates.filter((s) => new RegExp(`(^|[^0-9])0*${cno}([^0-9]|$)`).test(s.sheetName)))
      if (byName) return byName
    }
    return null
  }

  if (office.zone) {
    const zone = office.zone.trim().toLowerCase()
    const byLabel = withLabel.find(({ label }) => label && label.includes(zone) && !/^\d+\s*[-–]/.test(label))
    if (byLabel) return byLabel.sheet
    // A workbook covering several zones titles each zone's abstract tab with its zone ("Monitoring Format - 2025-26
    // (Kukatpally Zone)"), so pick the abstract that names this zone before guessing by position.
    const titled = sheets.find(
      (x) =>
        !/list/i.test(x.sheetName) &&
        !isCircleTabName(x.sheetName) &&
        /abstract/i.test(x.sheetName) &&
        x.grid.slice(0, 3).join(' ').toLowerCase().includes(zone)
    )
    if (titled) return titled
    // A zone-only office has no circle number. Its tab is often just an
    // abbreviation (e.g. "QBZ MF") that won't textually match the zone name,
    // so among the "MF" sheets that don't look like a per-circle sheet
    // ("C##"), fall back to the first one — the zone rollup is conventionally
    // the first sheet in these workbooks.
    const byName = candidates.filter((s) => !isCircleTabName(s.sheetName))
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

/** Tab names that spell a circle: "C57", "C-57", "C 57", "Circle 57", "C-12-Saroornagar". */
function isCircleTabName(name: string): boolean {
  return /^\s*c(?:ircle)?[\s.\-_#]*\d/i.test(name)
}

const isStatusHeader = (h: string) => h === 'status' || h === 'work status' || h.includes('current status')

/** A works-list sheet is recognised by its status column header, whatever the tab is called. */
function hasWorkStatusHeader(grid: string[][]): boolean {
  return grid.slice(0, 15).some((row) => row.some((c) => isStatusHeader((c || '').trim().toLowerCase())))
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
  const headerRowIdx = grid.findIndex(
    (row, i) => i < 15 && row.some((c) => isStatusHeader((c || '').trim().toLowerCase())) && row.some((c) => /name.*work/i.test((c || '').trim()))
  )
  if (headerRowIdx === -1) return []
  const header = grid[headerRowIdx]

  const col = {
    slNo: colIndex(header, (h) => /^sl\.?\s*no/.test(h)),
    circle: colIndex(header, (h) => h.includes('circle') || h === 'cir'),
    ward: colIndex(header, (h) => h.includes('ward')),
    workName: colIndex(header, (h) => /name.*work/.test(h)),
    estimateAmt: colIndex(header, (h) => h.includes('estimate')),
    typeOfWork: colIndex(header, (h) => (h.includes('type') && h.includes('work')) || h.startsWith('category')),
    sourceOfSanction: colIndex(header, (h) => h.includes('source')),
    sanctionDate: colIndex(header, (h) => h.includes('sanction') && h.includes('date')),
    agreementDate: colIndex(header, (h) => h.includes('agreement')),
    targetDate: colIndex(header, (h) => h.includes('target')),
    agencyDetails: colIndex(header, (h) => h.includes('agency')),
    // Some circles' sheets carry the real status in a plain "Status" column
    // and leave "Current status" blank on every row (the opposite of what
    // the header-row anchor above assumes) — try both, per row, and use
    // whichever one names a status (see the per-row pick below).
    statusCols: [
      colIndex(header, (h) => h === 'status' || h === 'work status'),
      colIndex(header, (h) => h.includes('current status'))
    ].filter(
      (i) => i >= 0
    ),
    winCode: colIndex(header, (h) => h.includes('win'))
  }

  // The works table ends at the Remarks column: anything the sheet adds after it is not taken.
  const remarksAt = header.findIndex((h) => (h || '').toLowerCase().includes('remark'))
  const keptCols = header.map((h, i) => ((h || '').trim() && (remarksAt < 0 || i <= remarksAt) ? i : -1)).filter((i) => i >= 0)
  const keptHeaders = keptCols.map((i) => (header[i] || '').replace(/\s+/g, ' ').trim())
  const rows: MonitoringFormatWorkRow[] = []
  for (let i = headerRowIdx + 1; i < grid.length; i++) {
    const row = grid[i]
    const slNo = cellAt(row, col.slNo)
    const workName = cellAt(row, col.workName)
    if (!slNo && !workName) continue
    // A banner line under the headings ("Circle 50- Madhapur" with only a ward filled in) has a title in the work-name
    // column but no serial number and no Win code: it is not a work.
    if (!slNo && !cellAt(row, col.winCode) && cellAt(row, col.workName).length < 40) continue
    // Some circles' sheets carry both a "Status" column (Progress / Completed …) and a
    // "Current status" column holding just the progress % ("upto 25%") — take the first
    // cell that actually names a status, not merely the first non-empty one.
    const status = col.statusCols.map((idx) => cellAt(row, idx)).find((s) => normalizeWorkStatus(s)) || ''
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
      winCode: cellAt(row, col.winCode),
      cells: keptCols.map((c) => (row[c] || '').trim()),
      ...(rows.length === 0 ? { headers: keptHeaders } : {})
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
  // Circle tabs are often just "C11-Nagole" / "C-12-Saroornagar" with no "list" in the name,
  // so a sheet also counts when it has a works-list status header and isn't an MF abstract.
  const candidates = sheets.filter(
    (s) => /list/i.test(s.sheetName) || (!/\bmf\b/i.test(s.sheetName) && hasWorkStatusHeader(s.grid))
  )
  if (candidates.length === 0) return []

  // A circle has several tabs ("C60 List", "C60 MF Spillover List AB" …); the one that really holds the works rows wins.
  const richest = (list: SheetGrid[]): SheetGrid[] => {
    let best: SheetGrid | null = null
    let bestCount = 0
    for (const s of list) {
      const count = parseMonitoringFormatWorkSheet(s).length
      if (count > bestCount) {
        best = s
        bestCount = count
      }
    }
    return best ? [best] : []
  }

  if (office.circle) {
    const cno = (office.circleNumber || '').trim()
    const circle = office.circle.trim().toLowerCase()
    const matches = candidates.filter((s) => {
      if (cno && new RegExp(`(^|[^0-9])0*${cno}([^0-9]|$)`).test(s.sheetName)) return true
      return circle ? s.sheetName.toLowerCase().includes(circle) : false
    })
    const match = richest(matches)
    if (match.length > 0) return match
    // The tab name doesn't carry the circle (e.g. just "List of works"): a lone list
    // sheet is the one, otherwise pick the sheet whose own Circle column names this circle.
    if (candidates.length === 1) return candidates
    const cellMatches = (v: string) => {
      const t = v.trim().toLowerCase()
      if (!t) return false
      if (circle && t.includes(circle)) return true
      return !!cno && new RegExp(`^0*${cno}([^0-9]|$)`).test(t)
    }
    let best: SheetGrid | null = null
    let bestCount = 0
    for (const s of candidates) {
      const count = parseMonitoringFormatWorkSheet(s).filter((w) => cellMatches(w.circle)).length
      if (count > bestCount) {
        best = s
        bestCount = count
      }
    }
    return best ? [best] : []
  }

  // A workbook holding every zone's data: keep only this zone's circles. Each circle's own list tab is complete,
  // so those are used (a zone-wide rollup tab would double-count the same works); with none found, fall back to
  // the zone's rollup list.
  if (office.zone && office.circleNumbers && office.circleNumbers.length > 0) {
    const perCircle = office.circleNumbers
      .map((cno) => richest(candidates.filter((s) => new RegExp(`^\\s*c[\\s.\\-_#]*0*${cno}([^0-9]|$)`, 'i').test(s.sheetName))))
      .flat()
    if (perCircle.length > 0) return perCircle
    const zone = office.zone.trim().toLowerCase()
    const inZone = candidates.filter((s) => s.grid.slice(0, 3).join(' ').toLowerCase().includes(zone))
    return inZone.length > 0 ? richest(inZone) : []
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

/** Strips a ward/circle pivot row's leading number (e.g. "276-Pragathi Nagar") so it can be matched against the work-list sheet's own ward/circle cell, which often carries just the name. */
function normalizeWardLabel(s: string): string {
  return s.trim().toLowerCase().replace(/^\d+\s*[-.]?\s*/, '')
}

/** Works matching one Ward Wise / Circle Wise tile — matched by name, not exact equality, since the pivot's row label and the work-list sheet's own ward/circle cell don't always carry the same number prefix. An optional status key further narrows to that ward's works in one status (Dashboard's Ward Wise "Status" filter). */
export function filterMonitoringFormatWorksByWard(
  works: MonitoringFormatWorkRow[],
  wardLabel: string,
  pivotKind: 'ward' | 'circle',
  statusKey?: MfStatusKey | null
): MonitoringFormatWorkRow[] {
  const target = normalizeWardLabel(wardLabel)
  if (!target) return []
  return works.filter((w) => {
    const value = normalizeWardLabel(pivotKind === 'circle' ? w.circle : w.ward)
    if (value === '' || !(value === target || target.includes(value) || value.includes(target))) return false
    return !statusKey || w.statusKey === statusKey
  })
}


/**
 * The works' columns exactly as the workbook lays them out: headings plus one string[] per work.
 * `allWorks` is the full imported list — the headings are stored only on the first row of each
 * sheet, so a status-filtered subset (which may not contain that row) must look them up there.
 * Sheets whose headings differ are merged into one column set, matched by heading text. Rows saved
 * before `cells` existed fall back to the fixed columns.
 */
const headerKey = (h: string): string => h.toLowerCase().replace(/\s+/g, ' ').trim()

/** monitoringFormatWorksTable with the user's deleted columns (matched by heading text) left out — used by the screen and every download alike. */
export function monitoringFormatWorksTable(
  works: MonitoringFormatWorkRow[],
  allWorks: MonitoringFormatWorkRow[] = works,
  hiddenHeaders: string[] = []
): ReturnType<typeof fullMonitoringFormatWorksTable> {
  const full = fullMonitoringFormatWorksTable(works, allWorks)
  if (hiddenHeaders.length === 0) return full
  const hidden = new Set(hiddenHeaders.map(headerKey))
  const keep = full.headers.map((h, c) => c).filter((c) => !hidden.has(headerKey(full.headers[c])))
  if (keep.length === full.headers.length) return full
  const indexOf = (old: number) => keep.indexOf(old)
  return {
    headers: keep.map((c) => full.headers[c]),
    rows: full.rows.map((r) => keep.map((c) => r[c])),
    slNoCol: indexOf(full.slNoCol),
    statusCol: indexOf(full.statusCol)
  }
}

function fullMonitoringFormatWorksTable(
  works: MonitoringFormatWorkRow[],
  allWorks: MonitoringFormatWorkRow[] = works
): {
  headers: string[]
  rows: string[][]
  /** Index of the Sl.No column (renumbered 1..n in downloads), or -1. */
  slNoCol: number
  /** Index of the status column, or -1. */
  statusCol: number
} {
  // Each work's sheet headings: the most recent `headers` seen walking the full list in order.
  const headersOf = new Map<MonitoringFormatWorkRow, string[]>()
  let current: string[] | undefined
  for (const w of allWorks) {
    if (w.headers) current = w.headers
    if (current) headersOf.set(w, current)
  }
  const key = (h: string) => h.toLowerCase().replace(/\s+/g, ' ').trim()
  if (works.length > 0 && works.every((w) => w.cells && headersOf.has(w))) {
    const headers: string[] = []
    const seen = new Map<string, number>()
    for (const w of works) {
      for (const h of headersOf.get(w)!) {
        if (!seen.has(key(h))) {
          seen.set(key(h), headers.length)
          headers.push(h)
        }
      }
    }
    const rows = works.map((w) => {
      const own = headersOf.get(w)!
      const row = headers.map(() => '')
      own.forEach((h, i) => {
        const at = seen.get(key(h))
        if (at !== undefined && !row[at]) row[at] = w.cells![i] ?? ''
      })
      return row
    })
    // Also applied here, so lists saved before the parser stopped at Remarks are cut the same way.
    const remarksCol = headers.findIndex((h) => h.toLowerCase().includes('remark'))
    const upTo = remarksCol >= 0 ? remarksCol + 1 : headers.length
    headers.length = upTo
    rows.forEach((r) => (r.length = upTo))
    const lower = headers.map((h) => h.toLowerCase())
    return {
      headers,
      rows,
      slNoCol: lower.findIndex((h) => /^sl\.?\s*no/.test(h)),
      statusCol: lower.findIndex((h) => h === 'status' || h === 'work status' || h.includes('current status'))
    }
  }
  return {
    headers: ['Sl.No', 'Circle', 'Ward', 'Name of Work', 'Type of Work', 'Win Code', 'Estimate (Rs. Lakhs)', 'Agency Details', 'Sanction Date', 'Agreement Date', 'Target Date', 'Status'],
    rows: works.map((w) => [w.slNo, w.circle, w.ward, w.workName, w.typeOfWork, w.winCode, String(w.estimateAmt), w.agencyDetails, w.sanctionDate, w.agreementDate, w.targetDate, w.status]),
    slNoCol: 0,
    statusCol: 11
  }
}

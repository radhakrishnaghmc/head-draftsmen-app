import { resolveColumns } from './columnMatch'
import type { ColumnEmbeddings, ColumnSpec } from './columnMatch'

export interface EstimateWorkItem {
  description: string
  quantity: string
  rate: string
  unit: string
  /** The estimate's own printed Amount for this item (its Amount-column cell on the summary row), when that column exists — used to catch items the estimate left un-costed (a quantity and rate, but a blank or zero Amount, so the estimate's own total silently omits them). Undefined when the estimate has no Amount column. */
  estimateAmount?: string
  /** Dimension breakdown (No's/L/B/D) that produced `quantity`, when the source estimate shows it — undefined when it doesn't (e.g. a hand-abstracted estimate with only the final Qty). */
  nos?: string
  l?: string
  b?: string
  d?: string
  /** 0-based grid position of the description cell, for writing back to the original file. */
  descRow?: number
  descCol?: number
  /** 0-based grid position of the rate cell, for writing back to the original file. */
  rateRow?: number
  rateCol?: number
  /**
   * True when this item is one of several measurement-row variants sharing a
   * single lead-row description (e.g. a depth-range breakdown) — its
   * `descRow`/`descCol` point at that row's own short label, not the shared
   * lead-row description, since the lead row can't be overwritten once per
   * variant without clobbering its siblings.
   */
  isVariant?: boolean
}

function norm(s: unknown): string {
  return String(s ?? '')
    .replace(/[\r\n]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

// Matches a "Name of Work" / "Name of the Work" label cell, capturing any
// text already inline after it (e.g. "Name of Work: Road from A to B").
const WORK_NAME_LABEL_RE = /^name\s+of\s+(?:the\s+)?work\b\s*:?\s*(.*)$/i

/**
 * Best-effort extraction of a "Label: value" style field from an estimate's
 * title block (the rows above the item header) — the label is usually
 * either followed inline by the value in the same cell (after a colon), or
 * sits alone with the value in the next non-empty cell along the same row.
 * Returns undefined if no such label is found — callers should treat that
 * as "couldn't tell", not "blank". `labelRe` must capture the inline
 * remainder (if any) in its first group, matching WORK_NAME_LABEL_RE's shape.
 */
export function extractLabeledField(grid: string[][], headerRowIndex: number, labelRe: RegExp): string | undefined {
  for (let r = 0; r < headerRowIndex; r++) {
    const row = grid[r] ?? []
    for (let c = 0; c < row.length; c++) {
      const cell = norm(row[c])
      if (!cell) continue
      const m = labelRe.exec(cell)
      if (!m) continue
      const inline = norm(m[1] ?? '')
      if (inline) return inline
      for (let c2 = c + 1; c2 < row.length; c2++) {
        const next = norm(row[c2])
        if (next) return next
      }
    }
  }
  return undefined
}

/**
 * Best-effort extraction of the work's name from the estimate's title block
 * (e.g. "Detailed Estimate for ..." / "Name of Work: ..."). See
 * extractLabeledField for the matching rules.
 */
export function extractWorkName(grid: string[][], headerRowIndex: number): string | undefined {
  return extractLabeledField(grid, headerRowIndex, WORK_NAME_LABEL_RE)
}

/**
 * Split one sheet's grid into separate estimate blocks when several complete
 * estimates are stacked in the same sheet — each estimate opens with its own
 * "Name of Work: ..." row, so every such row (after the first) starts a new
 * block that runs down to the row before the next one. Returns [grid]
 * unchanged when the sheet holds zero or one estimate, so callers can treat
 * the single- and multi-estimate cases the same way.
 */
export function splitEstimateBlocks(grid: string[][]): string[][][] {
  const marks: number[] = []
  grid.forEach((row, i) => {
    if (row.some((cell) => WORK_NAME_LABEL_RE.test(norm(cell)))) marks.push(i)
  })
  if (marks.length <= 1) return [grid]
  return marks.map((start, m) => grid.slice(start, m + 1 < marks.length ? marks[m + 1] : grid.length))
}

const GRAND_TOTAL_LABEL_RE = /grand\s*total/i

/**
 * The estimate's sanctioned Grand Total expressed in lakhs — the figure at the
 * very bottom of the estimate (after LC/seigniorage/GST/unforeseen), not the
 * ECV. Unlike extractEstimateAmountLakhs (which reads a labelled field in the
 * title block above the items), this scans the whole block for the "Grand
 * Total" row and takes the amount on it: the value in the Amount column when
 * that column is known, otherwise the first positive number on the row (a
 * trailing deviation figure like "0.00"/"-3979667" is skipped). Returns
 * undefined when there is no Grand Total row to read.
 */
export function extractGrandTotalLakhs(grid: string[][], headerRowIndex: number): number | undefined {
  const header = grid[headerRowIndex] ?? []
  const amountIdx = header.findIndex((c) => /\bamount\b/i.test(norm(c)))
  const asPositive = (v: unknown): number | undefined => {
    const n = Number(norm(v).replace(/,/g, ''))
    return Number.isFinite(n) && n > 0 ? n : undefined
  }
  for (const row of grid) {
    if (!row.some((c) => GRAND_TOTAL_LABEL_RE.test(norm(c)))) continue
    let val = amountIdx >= 0 ? asPositive(row[amountIdx]) : undefined
    if (val == null) {
      for (const c of row) {
        val = asPositive(c)
        if (val != null) break
      }
    }
    if (val != null) return Math.round((val / 100000) * 100) / 100
  }
  return undefined
}

// Serial number header varies ("S.No", "S.No.", "Sl.No", "Sl No", "Serial No"…);
// unit header reads "Unit"/"Units", "UOM", or "Per" across different templates.
export const ESTIMATE_COLUMN_SPECS: ColumnSpec[] = [
  { label: 'Serial Number', patterns: [/^sl?\.?\s*no\.?$|serial/i] },
  // A "Total Qty" column is the item's final quantity — preferred over an
  // intermediate "Qty per Day" / "Qty per unit" column that a day-rate estimate
  // also carries (and leaves blank on most rows), which the plain /qty/ pattern
  // would otherwise claim first for being further left, dropping every item
  // whose per-day cell is empty.
  { label: 'Quantity', patterns: [/total\s*(?:qty|quantity)/i, /qty|quantity/i] },
  { label: 'Rate', patterns: [/rate/i] },
  { label: 'Unit', patterns: [/unit|uom|^per$/i] }
]

// S.No / Qty / Rate are always labelled in these estimates; the Unit column
// sometimes isn't (see detectUnitColumnFromData), so it's resolved separately.
const ESTIMATE_REQUIRED_COLUMN_SPECS: ColumnSpec[] = ESTIMATE_COLUMN_SPECS.filter((s) => s.label !== 'Unit')

/**
 * Which of ESTIMATE_COLUMN_SPECS' columns (if any) only resolved via the
 * embedding fallback rather than a regex — worth a "please double-check
 * this match" notice in the UI. Returns [] both when nothing needed the
 * fallback and when column resolution failed outright (extractEstimateItems
 * itself is the source of truth for that failure).
 */
export function estimateColumnsMatchedViaEmbedding(
  grid: string[][],
  headerRowIndex: number,
  embeddings?: ColumnEmbeddings
): string[] {
  const header = (grid[headerRowIndex] ?? []).map(norm)
  try {
    return resolveColumns(header, ESTIMATE_COLUMN_SPECS, embeddings).viaEmbedding
  } catch {
    return []
  }
}

export interface ResolvedEstimateColumns {
  snoCol: number
  descCol: number
  qtyCol: number
  rateCol: number
  unitCol: number
  /** Dimension breakdown columns — best-effort, undefined when the sheet doesn't show them (e.g. only a final Qty). */
  nosCol?: number
  lCol?: number
  bCol?: number
  dCol?: number
  /** The Amount column — best-effort, undefined when the sheet has no labelled Amount column. */
  amountCol?: number
}

const NOS_HEADER_RE = /^no'?s\.?$/i
const L_HEADER_RE = /^l\.?$/i
const B_HEADER_RE = /^b\.?$/i
const D_HEADER_RE = /^d\.?$/i

/**
 * Best-effort resolution of the No's/L/B/D dimension columns from the header
 * row's own text — separate from resolveColumns/ESTIMATE_COLUMN_SPECS since
 * these are optional (many estimates only ever show the final Qty, not the
 * breakdown that produced it), so a missing one shouldn't fail the whole
 * extraction the way a missing Qty/Rate/Unit column does.
 */
export function resolveDimensionColumns(header: string[]): Pick<ResolvedEstimateColumns, 'nosCol' | 'lCol' | 'bCol' | 'dCol'> {
  const normalized = header.map(norm)
  const find = (re: RegExp): number | undefined => {
    const idx = normalized.findIndex((h) => re.test(h))
    return idx === -1 ? undefined : idx
  }
  return { nosCol: find(NOS_HEADER_RE), lCol: find(L_HEADER_RE), bCol: find(B_HEADER_RE), dCol: find(D_HEADER_RE) }
}

// A departmental rate-schedule code (BLD-CSTN-…, APSS/SS clause, TBSC/TBSP,
// ELEC, MORTH). Its presence marks a full item specification, as opposed to a
// bare measurement sub-label ("Footings", "Columns", "Roof Beams").
const SCHEDULE_CODE_RE = /BLD-|CSTN|APSS|TBS[CP]?|\bSS\s*-?\s*\d|ELEC|MORTH/i

// The same code anchored at the very start of a description — the tell that a
// row *begins a new item* even when the source estimate forgot to put a serial
// number on it (a real omission seen in the field: an item's spec cell is
// filled but its S.No cell was left blank, so it would otherwise be swallowed
// into the item above and mis-take that item's description).
const STARTS_WITH_SCHEDULE_CODE_RE = /^\s*(BLD-|TBS[CP]|APSS|MORTH|ELEC-|\bSS\s*-?\s*\d)/i

/**
 * Whether a lead-row description is a full item specification rather than a
 * bare sub-item label. Detailed estimates break one costed item (e.g. "Supply
 * and placing of M-25 Design Mix Concrete …") into lettered sub-parts a/b/c
 * ("Footings", "Pedastals", "Plinth Beams"), each on its own S.No'd row with
 * only that short label — the parent's full spec never repeats. We treat a
 * long description, or one carrying a rate-schedule code, as a full spec; a
 * short, code-less one as a sub-label that should inherit its parent's spec.
 */
function isFullSpecDescription(description: string): boolean {
  return description.length >= 40 || SCHEDULE_CODE_RE.test(description)
}

/**
 * The actual block-extraction algorithm, shared by both the automatic path
 * (extractEstimateItems, which resolves column indices from the header row's
 * own text) and the manual-override path (extractEstimateItemsWithColumns,
 * for when OCR has garbled the header row too badly to resolve automatically
 * and a person points at the columns instead).
 *
 * Each item spans a block of rows: a lead row carrying the S.No and, in the
 * very next cell, the description (often a merged cell with no header of its
 * own — so the description column is always positional, not name-matched),
 * one or more measurement rows, and a summary row carrying the final
 * Qty/Rate/Unit (searched for from the end of the block, since some items
 * have several measurement lines before the summary). Percentage add-on
 * rows (Labour cess, GST, TPQC, etc.) have an S.No and description but never
 * a Rate/Unit, so they resolve to nothing and are skipped automatically —
 * no special-casing needed.
 */
function extractEstimateItemsFromColumns(
  grid: string[][],
  headerRowIndex: number,
  { snoCol, descCol, qtyCol, rateCol, unitCol, nosCol, lCol, bCol, dCol, amountCol }: ResolvedEstimateColumns
): EstimateWorkItem[] {
  const items: EstimateWorkItem[] = []
  let block: { row: string[]; gridRow: number }[] = []
  // The full spec of the most recent "parent" item — a lead row that carries a
  // real specification but no measurement of its own because its quantity is
  // spread across lettered sub-parts on the rows below (a/b/c). Each sub-part's
  // bare label ("Footings") is prefixed with this so the emitted item keeps the
  // parent's spec; cleared once a self-contained full-spec item is emitted.
  let pendingParentDescription: string | undefined

  // A row's own No's/L/B/D cells if it has any, else undefined — used to find
  // the dimension row that precedes a variant's own measurement row.
  function readDims(row: string[]): Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined {
    const nos = nosCol !== undefined ? norm(row[nosCol]) : ''
    const l = lCol !== undefined ? norm(row[lCol]) : ''
    const b = bCol !== undefined ? norm(row[bCol]) : ''
    const d = dCol !== undefined ? norm(row[dCol]) : ''
    if (!nos && !l && !b && !d) return undefined
    return { nos: nos || undefined, l: l || undefined, b: b || undefined, d: d || undefined }
  }

  function resolveBlock() {
    if (block.length === 0) return
    const leadRow = block[0]
    const description = norm(leadRow.row[descCol])
    const measureIndices: number[] = []
    block.forEach(({ row }, idx) => {
      if (norm(row[rateCol]) !== '' && norm(row[unitCol]) !== '' && norm(row[qtyCol]) !== '') measureIndices.push(idx)
    })
    // A block with more than one measurement row is a single item broken
    // into several depth/size variants — e.g. "Drilling of tube wells…"
    // followed by "0 to 30 mtrs", "30 to 60 mtrs", "60 to 90 mtrs" rows,
    // each with its own Qty/Rate but sharing the lead row's description.
    // Emit one BOQ item per variant (tagging each with its own row label)
    // instead of collapsing them all into just the last measurement row.
    const multipleVariants = measureIndices.length > 1
    // A bare sub-item label (short, no rate-schedule code) inherits the parent
    // item's spec, when one is pending — see pendingParentDescription. Depth/
    // size variants keep their own lead description untouched.
    const inheritsParent = !multipleVariants && pendingParentDescription !== undefined && !isFullSpecDescription(description)
    const baseDescription = inheritsParent ? `${pendingParentDescription} - ${description}` : description
    measureIndices.forEach((idx, pos) => {
      const { row, gridRow } = block[idx]
      let rate = norm(row[rateCol])
      let unit = norm(row[unitCol])
      const quantity = norm(row[qtyCol])
      // Some sub-works print the summary line as "… Say | Qty | Cum | <rate> |
      // Amount", putting the unit token in the Rate column and the rate number
      // in the Unit column. Detect that inversion per row (unit reads numeric,
      // rate reads as a unit token) and swap them back so Rate stays numeric.
      if (looksNumeric(unit) && !looksNumeric(rate) && UNIT_RE.test(rate)) {
        const swap = rate
        rate = unit
        unit = swap
      }
      const label = norm(row[descCol])
      const itemDescription = multipleVariants && label && label !== description ? `${description}( ${label})` : baseDescription
      // Dimension data (if any) sits on this variant's own row, or on a row
      // between it and the previous variant's own measurement row — never
      // reaching back into an earlier variant's dimensions.
      const lowerBound = pos > 0 ? measureIndices[pos - 1] + 1 : 1
      let dims: Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined
      for (let k = idx; k >= lowerBound; k--) {
        dims = readDims(block[k].row)
        if (dims) break
      }
      if (itemDescription && quantity) {
        const estimateAmount = amountCol !== undefined ? norm(row[amountCol]) : undefined
        items.push({
          description: itemDescription,
          quantity,
          rate,
          unit,
          ...(estimateAmount !== undefined ? { estimateAmount } : {}),
          ...dims,
          // A variant's own row carries its short label in descCol — that's
          // what gets colored, since the shared lead row can't be uniquely
          // colored per variant. A non-variant item's description instead
          // lives on the lead row, which may be several rows above its rate.
          descRow: multipleVariants ? gridRow : leadRow.gridRow,
          descCol,
          rateRow: gridRow,
          rateCol,
          isVariant: multipleVariants
        })
      }
    })
    // A full-spec lead row that produced no measurement of its own is a parent
    // whose sub-parts follow; remember its spec for them. A full-spec row that
    // *did* emit is a self-contained item, ending any parent's sub-part run.
    if (isFullSpecDescription(description)) {
      pendingParentDescription = measureIndices.length === 0 ? description : undefined
    }
    block = []
  }

  for (let r = headerRowIndex + 1; r < grid.length; r++) {
    const row = grid[r] ?? []
    // Stop at the end of this estimate's item list — its grand "Total" row, the
    // "Add …%" abstract lines below it, or a repeated header where a *second*
    // estimate pasted into the same sheet begins. Without this, an unrelated
    // estimate stacked below (some CMC sheets carry two) would have its items
    // read into this work's Schedule A, so the total no longer matches the
    // estimate's own item total.
    if (isItemListEnd(row, snoCol, descCol, qtyCol)) break
    const sno = norm(row[snoCol])
    const desc = norm(row[descCol])
    // A new item opens on any serial-numbered, described row — or on a row
    // whose description itself starts with a rate-schedule code even without a
    // serial (a spec row the estimate left un-numbered), so it isn't folded
    // into the item above and given that item's description.
    const opensNewItem = (sno !== '' && desc !== '') || STARTS_WITH_SCHEDULE_CODE_RE.test(desc)
    if (opensNewItem) {
      resolveBlock() // close whatever block was open
      block = [{ row, gridRow: r }]
    } else if (block.length > 0) {
      block.push({ row, gridRow: r })
    }
  }
  resolveBlock()

  return items
}

/**
 * Whether `row` marks the end of an estimate's measured-item list: its grand
 * "Total" row, one of the "Add …%" abstract/overhead lines that follow it, or a
 * reprinted "S.No / Description of work" header where a second estimate stacked
 * into the same sheet starts. Item extraction stops here so the summary section
 * and any unrelated estimate below are never read as more items.
 */
function isItemListEnd(row: string[], snoCol: number, descCol: number, qtyCol: number): boolean {
  const desc = norm(row[descCol]).toLowerCase()
  const qty = norm(row[qtyCol]).toLowerCase()
  const sno = norm(row[snoCol]).toLowerCase()
  // A genuine total row tabulates its figure under the Quantity/Amount columns,
  // so a "Total" in the *Quantity* column reliably marks the end of the items.
  // A bare "Total" in the *description* column does not: detailed estimates use
  // it as an intra-item sub-total label (e.g. a flooring item whose measured
  // parts are summed under a "Total" row, with the real Qty/Rate several rows
  // below), and a multi-sub-work estimate repeats a sub-work "Total" between
  // sub-works — stopping at either would truncate the item list mid-estimate.
  // "Grand Total" is unambiguous wording never used as an intra-item label, so
  // it still ends the list even when it only labels the description column.
  const isTotal = /^grand\s+total$/.test(desc) || /^(grand\s+)?total$/.test(qty)
  const isAbstractAdd =
    /^add\b/.test(desc) && /(charge|gst|\bvat\b|cess|seign[io]rage|\bnac\b|dmft|smet|permit|royalty|labour|contingenc|u\/f items|ls for)/.test(desc)
  const isNewEstimateHeader = /^s\.?\s*no\.?$/.test(sno) && /descrip/.test(desc)
  return isTotal || isAbstractAdd || isNewEstimateHeader
}

function looksNumeric(s: string): boolean {
  const t = s.trim().replace(/,/g, '')
  return t !== '' && Number.isFinite(Number(t))
}

/**
 * Some departmental templates squeeze a "Per" multiplier column ("Rate ...
 * Per 1 ... Cum" — almost always just the bare number 1) directly before the
 * real unit-text column, and leave that real unit column's own header cell
 * blank — so `/unit|^per$/i` (ESTIMATE_COLUMN_SPECS) claims the "Per" column
 * as "Unit" and every item ends up with "1" as its unit instead of "Cum"/
 * "Sqm"/etc.
 *
 * Detected after the fact, from the data rows themselves rather than the
 * header: if the resolved Unit column is overwhelmingly bare numbers across
 * the item rows while the very next column instead holds recognized unit
 * tokens, that next column is almost certainly the real one.
 */
/**
 * Locate the Unit column from the *data* when the header row doesn't label it
 * — a common detailed-estimate layout keeps an unlabelled column between Rate
 * and Amount that carries the unit token (Cum/Sqm/…) only on each item's
 * summary line. Picks the unclaimed column with the most recognized unit
 * tokens across the item rows, breaking ties toward the column just after
 * Rate (where these templates place it). Returns -1 if no column looks like a
 * unit column at all.
 */
function detectUnitColumnFromData(
  grid: string[][],
  headerRowIndex: number,
  rateCol: number,
  claimed: Set<number>
): number {
  let width = 0
  for (let r = headerRowIndex; r < grid.length; r++) width = Math.max(width, grid[r]?.length ?? 0)
  let bestCol = -1
  let bestHits = 0
  let bestDist = Infinity
  for (let c = 0; c < width; c++) {
    if (claimed.has(c)) continue
    let hits = 0
    for (let r = headerRowIndex + 1; r < grid.length; r++) {
      const cell = norm(grid[r]?.[c] ?? '')
      if (cell && UNIT_RE.test(cell)) hits++
    }
    const dist = Math.abs(c - (rateCol + 1))
    if (hits > bestHits || (hits === bestHits && hits > 0 && dist < bestDist)) {
      bestHits = hits
      bestCol = c
      bestDist = dist
    }
  }
  return bestHits > 0 ? bestCol : -1
}

function fixUnitColumnIfNumeric(grid: string[][], headerRowIndex: number, unitCol: number): number {
  const candidate = unitCol + 1
  let numeric = 0
  let unitLike = 0
  let sampled = 0
  for (let r = headerRowIndex + 1; r < grid.length && sampled < 30; r++) {
    const cell = norm(grid[r]?.[unitCol])
    if (!cell) continue
    sampled++
    if (looksNumeric(cell)) numeric++
    if (UNIT_RE.test(norm(grid[r]?.[candidate] ?? ''))) unitLike++
  }
  return sampled > 0 && numeric / sampled > 0.6 && unitLike > 0 ? candidate : unitCol
}

/**
 * Parse a "detailed abstract estimate" sheet into its real work items,
 * resolving the S.No/Qty/Rate/Unit columns from the header row's own text.
 *
 * Column headers are matched by regex first (fast, no model needed); when
 * `embeddings` is supplied and a header doesn't match any pattern, the
 * closest-matching header by semantic similarity is used instead — see
 * core/columnMatch.ts.
 */
export function extractEstimateItems(
  grid: string[][],
  headerRowIndex: number,
  embeddings?: ColumnEmbeddings
): EstimateWorkItem[] {
  const header = (grid[headerRowIndex] ?? []).map(norm)
  let snoCol: number, qtyCol: number, rateCol: number, unitCol: number
  try {
    // Fast path: a fully-labelled estimate (Serial/Qty/Rate/Unit headers).
    const resolved = resolveColumns(header, ESTIMATE_COLUMN_SPECS, embeddings)
    snoCol = resolved.indexByLabel['Serial Number']
    qtyCol = resolved.indexByLabel['Quantity']
    rateCol = resolved.indexByLabel['Rate']
    unitCol = resolved.indexByLabel['Unit']
  } catch {
    // Fallback: estimates that don't label the Unit column but still carry the
    // unit in an unlabelled column (Cum/Sqm on each item's summary line).
    // Resolve the three always-labelled columns, then find the unit from data.
    let required
    try {
      required = resolveColumns(header, ESTIMATE_REQUIRED_COLUMN_SPECS, embeddings)
    } catch {
      throw new Error('Could not find S.No / Qty / Rate / Unit columns in the estimate.')
    }
    snoCol = required.indexByLabel['Serial Number']
    qtyCol = required.indexByLabel['Quantity']
    rateCol = required.indexByLabel['Rate']
    // Exclude the already-identified columns (S.No, its Description at snoCol+1,
    // Qty, Rate) so those can't be mistaken for the unit column.
    const claimed = new Set([snoCol, snoCol + 1, qtyCol, rateCol])
    unitCol = detectUnitColumnFromData(grid, headerRowIndex, rateCol, claimed)
    if (unitCol === -1) {
      throw new Error('Could not find S.No / Qty / Rate / Unit columns in the estimate.')
    }
  }
  unitCol = fixUnitColumnIfNumeric(grid, headerRowIndex, unitCol)
  // Description isn't reliably labelled (and is often a merged cell with no
  // header of its own) — it's always the cell right after the serial number.
  const descCol = snoCol + 1
  if (descCol >= header.length) {
    throw new Error('Could not find a description column next to the S.No column in the estimate.')
  }
  const dims = resolveDimensionColumns(header)
  // The Amount column, when labelled — used only to catch un-costed items (a
  // quantity and rate but a blank/zero Amount). Kept distinct from the unit
  // column already claimed, and never the unlabelled formula column, so a
  // labelled "Amount"/"Cost" header is required. Optional: absent → not read.
  const amountIdx = header.findIndex((h, i) => i !== unitCol && /\bamount\b|\bcost\b/i.test(h))
  const amountCol = amountIdx === -1 ? undefined : amountIdx
  return extractEstimateItemsFromColumns(grid, headerRowIndex, { snoCol, descCol, qtyCol, rateCol, unitCol, amountCol, ...dims })
}

/**
 * The extracted items the estimate itself left un-costed: a real quantity and
 * rate (so Qty × Rate is a positive amount) but a blank or zero Amount cell in
 * the estimate's own Amount column. The estimate's printed total silently
 * omits these, so a BOQ built as Σ(Qty × Rate) will exceed the estimate's own
 * figure by exactly their worth — worth surfacing so it's a deliberate choice,
 * not a surprise. Returns [] when the estimate has no Amount column to compare
 * against (nothing can be judged missing), or when every item is costed.
 */
export function itemsMissingEstimateAmount(items: EstimateWorkItem[]): EstimateWorkItem[] {
  const num = (s: string | undefined): number => Number(String(s ?? '').replace(/,/g, ''))
  return items.filter((it) => {
    if (it.estimateAmount === undefined) return false // no Amount column at all
    const computed = Math.round((num(it.quantity) || 0) * (num(it.rate) || 0))
    if (computed <= 0) return false // genuinely zero work (e.g. a 0-qty provisional line)
    const printed = num(it.estimateAmount)
    return !Number.isFinite(printed) || Math.round(printed) === 0
  })
}

/**
 * Same block-extraction algorithm as extractEstimateItems, but with the
 * column indices pointed at directly instead of resolved from header text —
 * the fallback for a photographed estimate whose header row OCR'd too
 * poorly (garbled by a skewed photo, small dense header font, etc.) for
 * automatic resolution to work at all. Used when a person manually reviews
 * the raw OCR grid and marks which row is the header and which column is
 * which.
 */
export function extractEstimateItemsWithColumns(
  grid: string[][],
  headerRowIndex: number,
  columns: ResolvedEstimateColumns
): EstimateWorkItem[] {
  return extractEstimateItemsFromColumns(grid, headerRowIndex, columns)
}

// Unit abbreviations used across the department's rate schedules (APSS/SS)
// for measuring work items — the one reliable anchor for finding an item's
// summary line directly in OCR'd text, since it's short, distinctive, and
// (unlike column position) survives a table-line OCR engine returning whole
// lines of text instead of per-word/per-cell boxes. Beyond the civil-works
// units, service/transport/labour estimates measure by Tonne, Trip, Day,
// Hour, Km, Pair etc. — included so those estimates' item lines are found too
// (a garbage-transport estimate priced "…/tonne" is a real case). Longer
// spellings precede their abbreviations so the regex prefers the fuller match.
const UNIT_TOKENS = [
  'Cum',
  'Sqm',
  'Rmt',
  'Nos',
  'Kg',
  'MT',
  'Ltr',
  'Mtr',
  'Sft',
  'RM',
  'Each',
  'Tonnes',
  'Tonne',
  'Ton',
  'Trip',
  'Days',
  'Day',
  'Hour',
  'Hrs',
  'Hr',
  'Kms',
  'Km',
  'Cft',
  'Brass',
  'Pair',
  'Litre'
]
// No leading \b: OCR frequently glues the "Per" multiplier straight onto the
// unit with no space (e.g. "2345.001Cum"), which would otherwise fail a
// word-boundary check between the digit and the letter. Exported so the
// Schedule A builder can detect an unlabelled unit column the same way.
export const UNIT_RE = new RegExp(`(${UNIT_TOKENS.join('|')})\\b`, 'i')
const NUMBER_RE = /\d[\d,]*\.\d{2}/g

/**
 * A line that is only numbers (with at most a stray "x" multiplier / column
 * separators) and no unit token or real words — a summary row's Qty/Rate (or
 * Amount) cell that a line-detecting OCR split off as its own line. Used to
 * re-stitch a summary row the OCR fragmented across two lines (e.g. "tonne
 * 5577686.00" on one line and "14960.00 372.84" on the next).
 */
function looksNumericFragment(text: string): boolean {
  if (UNIT_RE.test(text)) return false
  if (!(text.match(NUMBER_RE)?.length ?? 0)) return false
  return text.replace(/[0-9.,/\-\s]/g, '').replace(/x/gi, '') === ''
}

/** How closely Qty × Rate must reproduce Amount to count as consistent (1%, covering display rounding). */
function amountTolerance(amount: number): number {
  return Math.max(1, amount * 0.01)
}

/**
 * From the decimal numbers on an item's summary line, choose which are Qty,
 * Rate and Amount. The printed order is "Qty | Rate | Per | Amount", so the
 * naive choice is the last three. But OCR routinely leaves an *extra* decimal
 * number on the line — a dimension figure, a running page total, a stray "Per"
 * multiplier — which would shift that window. So we lean on the one invariant a
 * detailed estimate always holds: Amount = Qty × Rate. We look for the
 * right-most Qty/Rate/Amount triple whose product matches, which lands on the
 * true columns even amid extra numbers. Only when no triple is consistent do we
 * fall back to the printed last-three. Values with no 2-decimal form (a
 * whole-number No's) never appear in `numbers`, so can't be mistaken for Qty.
 */
function pickQtyRateAmount(numbers: string[]): { quantity: string; rate: string; amount?: string } {
  const n = numbers.map((s) => Number(s.replace(/,/g, '')))
  for (let a = n.length - 1; a >= 2; a--) {
    if (!(n[a] > 0)) continue
    const tol = amountTolerance(n[a])
    for (let r = a - 1; r >= 1; r--) {
      for (let q = r - 1; q >= 0; q--) {
        if (n[q] > 0 && n[r] > 0 && Math.abs(n[q] * n[r] - n[a]) <= tol) {
          return { quantity: numbers[q], rate: numbers[r], amount: numbers[a] }
        }
      }
    }
  }
  const [quantity, rate, amount] = numbers.slice(-3)
  return { quantity, rate, amount }
}

/**
 * Correct a misread quantity against the authoritative Amount = Qty × Rate.
 * When a summary line carries an Amount (the estimate's own printed total) that
 * the read Qty × Rate can't reproduce, and Amount ÷ Rate is a clean positive
 * value, we trust the Amount and recompute the quantity: OCR trips on the Qty
 * column's digits far more often than on the (wider, more redundant) Amount, so
 * this keeps the item's amount — and the grand total, which the app re-derives
 * as Σ(Qty × Rate) — faithful to the source. Returns the original quantity
 * untouched when the row is already consistent, has no usable Amount/Rate, or
 * the recomputation isn't finite.
 */
function reconcileQuantity(quantity: string, rate: string, amount: string | undefined): string {
  const q = Number(quantity.replace(/,/g, ''))
  const r = Number(rate.replace(/,/g, ''))
  const a = amount != null ? Number(amount.replace(/,/g, '')) : NaN
  if (!(r > 0) || !Number.isFinite(a) || !(a > 0)) return quantity
  if (Math.abs(q * r - a) <= amountTolerance(a)) return quantity
  const corrected = a / r
  return Number.isFinite(corrected) && corrected > 0 ? corrected.toFixed(2) : quantity
}

// A bare-integer No's count ("1", "2", "10") — as printed in the No's column,
// with no decimal point (distinct from a measurement).
const DIM_COUNT_RE = /^\d{1,4}$/
// An L/B/D measurement — always carries a decimal point, but any number of
// decimals (a thickness is often 3, e.g. "0.075"), unlike a Qty/Rate/Amount
// figure which the summary line pins to exactly 2.
const DIM_MEASURE_RE = /^\d[\d,]*\.\d+$/

/**
 * A "No's / L / B / D" dimension line has no unit token and no other text —
 * just the count and 2 to 3 measurements, printed in that column order.
 * Distinguished from a description line (which has actual words, so never
 * matches purely numeric tokens) and from a summary line (which always carries
 * a unit token).
 *
 * The No's is usually a *bare integer* ("1 250.00 6.00 0.15") — so it's matched
 * separately from the decimal measurements rather than demanding it too be a
 * 2-decimal number, which used to reject the whole (perfectly clean) line and
 * push it into the far more error-prone glued-number factoring below. When
 * there's no leading count, 4 numbers is No's+L+B+D, 3 is L+B+D, and 2 is L+B
 * (an area item).
 */
function parseDimensionLine(text: string): Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined {
  const tokens = text.split(/\s+/)
  if (tokens.length < 2 || tokens.length > 4) return undefined

  // A leading bare-integer count, then all-measurement dimensions.
  if (DIM_COUNT_RE.test(tokens[0])) {
    const [nos, ...dims] = tokens
    if (dims.length < 2 || dims.length > 3 || !dims.every((t) => DIM_MEASURE_RE.test(t))) return undefined
    const [l, b, d] = dims
    return dims.length === 3 ? { nos, l, b, d } : { nos, l, b }
  }

  // No leading count — every token is a measurement.
  if (!tokens.every((t) => DIM_MEASURE_RE.test(t))) return undefined
  if (tokens.length === 4) {
    const [nos, l, b, d] = tokens
    return { nos, l, b, d }
  }
  if (tokens.length === 3) {
    const [l, b, d] = tokens
    return { l, b, d }
  }
  const [l, b] = tokens
  return { l, b }
}

// The final Qty print is trusted (rounding to 2 decimals from a real
// multiplication) — a reconstructed No's×L×B×D product is accepted only
// within floating-point/rounding noise of it, not a loose approximation. A
// genuine OCR misread of a dimension digit throws the product off by far
// more than this, so it correctly fails closed rather than accepting a
// wrong number.
const DIMENSION_PRODUCT_TOLERANCE = 0.02

/**
 * Split a run of digits-and-dots with no delimiters (e.g. "1550.000.900.20")
 * into consecutive `digits.digits` tokens. With a single dot there's only
 * one number here at all — whatever follows the dot is its decimal part,
 * however many digits (a thickness like "0.075" is common and shouldn't be
 * rejected just for having 3). With 2+ dots there's real ambiguity about
 * where one token's decimals end and the next's integer part begins, so
 * each dot is deterministically treated as *some* token's own decimal
 * point with exactly the next 2 digits as its decimal part (the department's
 * usual convention) and whatever sits between that and the next dot as the
 * following token's integer part. Returns undefined if the string doesn't
 * fully consume this way — i.e. it isn't actually shaped like a
 * concatenation of `\d+\.\d\d` tokens (or, for a single dot, `\d+\.\d+`) at
 * all.
 */
function splitGluedDecimalTokens(s: string): string[] | undefined {
  const dots: number[] = []
  for (let i = 0; i < s.length; i++) if (s[i] === '.') dots.push(i)
  if (dots.length === 0) return undefined
  if (dots.length === 1) {
    const intPart = s.slice(0, dots[0])
    const decPart = s.slice(dots[0] + 1)
    return /^\d+$/.test(intPart) && /^\d+$/.test(decPart) ? [`${intPart}.${decPart}`] : undefined
  }
  const tokens: string[] = []
  let pos = 0
  for (const dot of dots) {
    const intPart = s.slice(pos, dot)
    const decPart = s.slice(dot + 1, dot + 3)
    if (!/^\d+$/.test(intPart) || !/^\d{2}$/.test(decPart)) return undefined
    tokens.push(`${intPart}.${decPart}`)
    pos = dot + 3
  }
  return pos === s.length ? tokens : undefined
}

function mapPiecesToDims(pieces: string[], hasNos: boolean): Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined {
  if (hasNos) {
    if (pieces.length === 4) return { nos: pieces[0], l: pieces[1], b: pieces[2], d: pieces[3] }
    if (pieces.length === 3) return { nos: pieces[0], l: pieces[1], b: pieces[2] }
    if (pieces.length === 2) return { nos: pieces[0], l: pieces[1] }
  } else {
    if (pieces.length === 4) return { nos: pieces[0], l: pieces[1], b: pieces[2], d: pieces[3] }
    if (pieces.length === 3) return { l: pieces[0], b: pieces[1], d: pieces[2] }
    if (pieces.length === 2) return { l: pieces[0], b: pieces[1] }
  }
  return undefined
}

const MAX_DIMENSION_PIECES = 4

/**
 * Try to recover No's/L/B/D starting from `runs[startIdx]` (already isolated
 * from surrounding words — see factorGluedDimensions), against the item's
 * own already-trusted Qty. Handles two distinct ways OCR has been observed
 * to glue this row's numbers together:
 *
 * - all of them fused into runs[startIdx] itself with no delimiter at all
 *   (e.g. "1550.000.900.20" for No's=1/L=550.00/B=0.90/D=0.20) — recovered
 *   by peeling 0-3 leading digits off the front as a bare-integer No's and
 *   deterministically splitting the remainder into decimal tokens
 *   (splitGluedDecimalTokens);
 * - only some of them fused, with the rest already clean, separate tokens
 *   later on the same line (e.g. "1100.006.00 0.30" for No's=1/L=100.00/
 *   B=6.00 glued, then D=0.30 on its own) — recovered by extending the
 *   piece list forward into subsequent runs' own (unpeeled) tokens.
 *
 * Tries every peel first, then — per peel — every resulting piece count
 * from longest (4) down to 2, always anchored at the start of the sequence
 * (real No's/L/B/D always leads; anything after is Qty/Amount noise), and
 * accepts the first whose product matches Qty. A lone single-piece result
 * is never accepted even if it numerically equals Qty: with only one piece
 * there's no way to tell genuine dimension data apart from Qty/Amount just
 * being printed again on the same line, so accepting it would risk
 * reporting a coincidence as fact.
 */
function tryFactorFrom(runs: string[], startIdx: number, targetQty: number): Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined {
  const run = runs[startIdx]
  for (let peel = 0; peel <= 3 && peel < run.length; peel++) {
    const nos = peel > 0 ? run.slice(0, peel) : undefined
    const ownTokens = splitGluedDecimalTokens(run.slice(peel))
    if (!ownTokens) continue

    const pieces = [...(nos ? [nos] : []), ...ownTokens]
    for (let j = startIdx + 1; j < runs.length && pieces.length < MAX_DIMENSION_PIECES; j++) {
      const extra = splitGluedDecimalTokens(runs[j])
      if (!extra) break // a later run that isn't itself clean decimal tokens ends the dimension sequence
      pieces.push(...extra)
    }

    for (let len = Math.min(MAX_DIMENSION_PIECES, pieces.length); len >= 2; len--) {
      const window = pieces.slice(0, len)
      const product = window.reduce((p, t) => p * Number(t), 1)
      if (Math.abs(product - targetQty) > DIMENSION_PRODUCT_TOLERANCE) continue
      const dims = mapPiecesToDims(window, Boolean(nos))
      if (dims) return dims
    }
  }
  return undefined
}

/**
 * Recover No's/L/B/D from a line where OCR glued the dimension numbers
 * together with no delimiter, possibly alongside real description words on
 * the same line (e.g. "for UGD line X 1550.000.900.20 99.0026642 2616300") —
 * the shape a line-detecting OCR engine (see electron/ocr.ts) actually
 * produces for this row when the printed columns sit close together,
 * unlike the clean, evenly-spaced case parseDimensionLine already handles.
 * Isolates each maximal digit-and-dot run in the line and tries factoring
 * starting from each one in turn (tryFactorFrom) against the item's own
 * already-trusted Qty, taking the first that resolves unambiguously.
 * Returns undefined (leave the dimension row blank) rather than ever guess —
 * better than silently printing a wrong figure onto an official estimate.
 */
function factorGluedDimensions(text: string, targetQty: number): Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined {
  const runs = text.match(/[0-9.]+/g) ?? []
  for (let i = 0; i < runs.length; i++) {
    const found = tryFactorFrom(runs, i, targetQty)
    if (found) return found
  }
  return undefined
}

/** Tries factorGluedDimensions against each of an item's accumulated description lines in order, taking the first that resolves. */
function findGluedDimensions(
  descriptionLines: string[],
  targetQty: number
): Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined {
  for (const line of descriptionLines) {
    const found = factorGluedDimensions(line, targetQty)
    if (found) return found
  }
  return undefined
}

/**
 * Parse a "detailed abstract estimate" directly from OCR'd *line* text (one
 * string per detected line, in top-to-bottom reading order) — no column
 * positions involved at all. This is the extraction path for a
 * line-detecting OCR engine (see electron/ocr.ts): its detector already
 * returns each printed line as one clean, correctly-spelled unit of text
 * (unlike a per-word/per-character OCR engine, whose word boxes have to be
 * re-assembled into rows and columns by position — the reconstruction that
 * core/ocrTableReconstruct.ts does, and that this estimate's own dense,
 * merged-description-cell layout defeats almost every time).
 *
 * Instead of hunting for a header row and mapping columns, this walks the
 * lines looking for each item's *summary* line — recognized by a unit
 * abbreviation (UNIT_TOKENS) plus at least 3 decimal numbers on the same
 * line — and takes the last 3 as Qty/Rate/Amount, matching the printed
 * "Qty | Rate | Per | Amount" column order. Every other line accumulates as
 * that item's description; a summary line closes the block and starts the
 * next one. Lines before the printed "Qty ... Rate ..." header (the title
 * block, corporation letterhead, "Name of Work" line) are skipped so they
 * never leak into the first item's description. "Qty" and "Rate" don't
 * always land on the same detected line — a tall/wrapped header row can get
 * split across two OCR lines (e.g. "SI. Rate Per Amount" then "Description
 * of work No's L B D Qty") — so each is tracked independently and the
 * header counts as done once both have been seen, not only when one line
 * contains both.
 */
export function extractEstimateItemsFromLines(lines: string[]): EstimateWorkItem[] {
  const items: EstimateWorkItem[] = []
  let description: string[] = []
  let pendingDims: Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined
  let pastHeader = false
  let sawQty = false
  let sawRate = false

  for (let i = 0; i < lines.length; i++) {
    const text = norm(lines[i])
    if (!text) continue

    if (!pastHeader) {
      if (/\bqty\b/i.test(text)) sawQty = true
      if (/\brate\b/i.test(text)) sawRate = true
      if (sawQty && sawRate) pastHeader = true
      continue
    }

    const dims = parseDimensionLine(text)
    if (dims) {
      pendingDims = dims
      continue
    }

    const unitMatch = UNIT_RE.exec(text)
    let numbers: string[] = text.match(NUMBER_RE) ?? []
    // A unit-bearing line with too few numbers is a summary row the OCR split
    // across two lines — its unit+Amount cell detached from its Qty/Rate cell
    // (and the two can even land out of reading order). Re-stitch by pulling
    // the Qty/Rate figures from an adjacent numeric-only fragment (the line
    // just accumulated, or the next line), always ordered Qty/Rate first and
    // this line's Amount last so pickQtyRateAmount still reads the columns right.
    let consumedNext = false
    if (unitMatch && numbers.length < 3) {
      const prev = description[description.length - 1]
      const next = i + 1 < lines.length ? norm(lines[i + 1]) : ''
      let frag: string[] = []
      let poppedPrev = false
      if (prev && looksNumericFragment(prev)) {
        frag = prev.match(NUMBER_RE) ?? []
        poppedPrev = true
      }
      if (frag.length + numbers.length < 3 && next && looksNumericFragment(next)) {
        frag = [...frag, ...(next.match(NUMBER_RE) ?? [])]
        consumedNext = true
      }
      if (frag.length + numbers.length >= 3) {
        numbers = [...frag, ...numbers]
        if (poppedPrev) description.pop()
      } else {
        consumedNext = false
      }
    }
    if (unitMatch && numbers.length >= 3 && description.join(' ').trim()) {
      const picked = pickQtyRateAmount(numbers)
      const rate = picked.rate
      const quantity = reconcileQuantity(picked.quantity, rate, picked.amount)
      // The clean, cleanly-spaced dimension line (parseDimensionLine, above)
      // takes priority when present; otherwise fall back to reconstructing
      // No's/L/B/D out of whichever accumulated line has them glued into a
      // noisy, word-mixed blob, now that Qty (the reconstruction's own
      // check) is finally known.
      const dims = pendingDims ?? findGluedDimensions(description, Number(quantity))
      items.push({ description: description.join(' ').trim(), quantity, rate, unit: unitMatch[1], ...dims })
      description = []
      pendingDims = undefined
      if (consumedNext) i++ // the next line's figures were folded into this summary
      continue
    }

    description.push(text)
  }

  return items
}

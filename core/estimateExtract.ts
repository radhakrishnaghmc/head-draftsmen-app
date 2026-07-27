import { resolveColumns } from './columnMatch'
import type { ColumnEmbeddings, ColumnSpec } from './columnMatch'

export interface EstimateWorkItem {
  description: string
  quantity: string
  rate: string
  unit: string
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

// Serial number header varies ("S.No", "S.No.", "Sl.No", "Sl No", "Serial No"…);
// unit header sometimes reads "Per" instead of "Unit".
export const ESTIMATE_COLUMN_SPECS: ColumnSpec[] = [
  { label: 'Serial Number', patterns: [/^sl?\.?\s*no\.?$|serial/i] },
  { label: 'Quantity', patterns: [/qty|quantity/i] },
  { label: 'Rate', patterns: [/rate/i] },
  { label: 'Unit', patterns: [/unit|^per$/i] }
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
  { snoCol, descCol, qtyCol, rateCol, unitCol, nosCol, lCol, bCol, dCol }: ResolvedEstimateColumns
): EstimateWorkItem[] {
  const items: EstimateWorkItem[] = []
  let block: { row: string[]; gridRow: number }[] = []

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
    measureIndices.forEach((idx, pos) => {
      const { row, gridRow } = block[idx]
      const rate = norm(row[rateCol])
      const unit = norm(row[unitCol])
      const quantity = norm(row[qtyCol])
      const label = norm(row[descCol])
      const itemDescription = multipleVariants && label && label !== description ? `${description}( ${label})` : description
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
        items.push({
          description: itemDescription,
          quantity,
          rate,
          unit,
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
    block = []
  }

  for (let r = headerRowIndex + 1; r < grid.length; r++) {
    const row = grid[r] ?? []
    const sno = norm(row[snoCol])
    const desc = norm(row[descCol])
    if (sno !== '' && desc !== '') {
      resolveBlock() // close whatever block was open
      block = [{ row, gridRow: r }]
    } else if (block.length > 0) {
      block.push({ row, gridRow: r })
    }
  }
  resolveBlock()

  return items
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
  return extractEstimateItemsFromColumns(grid, headerRowIndex, { snoCol, descCol, qtyCol, rateCol, unitCol, ...dims })
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
// lines of text instead of per-word/per-cell boxes.
const UNIT_TOKENS = ['Cum', 'Sqm', 'Rmt', 'Nos', 'Kg', 'MT', 'Ltr', 'Mtr', 'Sft', 'RM', 'Each']
// No leading \b: OCR frequently glues the "Per" multiplier straight onto the
// unit with no space (e.g. "2345.001Cum"), which would otherwise fail a
// word-boundary check between the digit and the letter.
const UNIT_RE = new RegExp(`(${UNIT_TOKENS.join('|')})\\b`, 'i')
const NUMBER_RE = /\d[\d,]*\.\d{2}/g
const NUMBER_TOKEN_RE = /^\d[\d,]*\.\d{2}$/

/**
 * A "No's / L / B / D" dimension line has no unit token and no other text —
 * just 2 to 4 decimal numbers, printed in that column order. Distinguished
 * from a description line (which has actual words, so never matches purely
 * numeric tokens) and from a summary line (which always carries a unit
 * token). 4 numbers is the unambiguous No's+L+B+D case; 3 is L+B+D with an
 * implicit single count; 2 is treated as L+B (e.g. an area item), the most
 * common two-number case.
 */
function parseDimensionLine(text: string): Pick<EstimateWorkItem, 'nos' | 'l' | 'b' | 'd'> | undefined {
  const tokens = text.split(/\s+/)
  if (tokens.length < 2 || tokens.length > 4 || !tokens.every((t) => NUMBER_TOKEN_RE.test(t))) return undefined
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

  for (const raw of lines) {
    const text = norm(raw)
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
    const numbers = text.match(NUMBER_RE) ?? []
    if (unitMatch && numbers.length >= 3 && description.join(' ').trim()) {
      const [quantity, rate] = numbers.slice(-3)
      // The clean, cleanly-spaced dimension line (parseDimensionLine, above)
      // takes priority when present; otherwise fall back to reconstructing
      // No's/L/B/D out of whichever accumulated line has them glued into a
      // noisy, word-mixed blob, now that Qty (the reconstruction's own
      // check) is finally known.
      const dims = pendingDims ?? findGluedDimensions(description, Number(quantity))
      items.push({ description: description.join(' ').trim(), quantity, rate, unit: unitMatch[1], ...dims })
      description = []
      pendingDims = undefined
      continue
    }

    description.push(text)
  }

  return items
}

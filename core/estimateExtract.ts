import { resolveColumns } from './columnMatch'
import type { ColumnEmbeddings, ColumnSpec } from './columnMatch'

export interface EstimateWorkItem {
  description: string
  quantity: string
  rate: string
  unit: string
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

/**
 * Parse a "detailed abstract estimate" sheet into its real work items. Each
 * item spans a block of rows: a lead row carrying the S.No and, in the very
 * next cell, the description (often a merged cell with no header of its
 * own — so the description column is always positional, not name-matched),
 * one or more measurement rows, and a summary row carrying the final
 * Qty/Rate/Unit (searched for from the end of the block, since some items
 * have several measurement lines before the summary). Percentage add-on
 * rows (Labour cess, GST, TPQC, etc.) have an S.No and description but never
 * a Rate/Unit, so they resolve to nothing and are skipped automatically —
 * no special-casing needed.
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
    const resolved = resolveColumns(header, ESTIMATE_COLUMN_SPECS, embeddings)
    snoCol = resolved.indexByLabel['Serial Number']
    qtyCol = resolved.indexByLabel['Quantity']
    rateCol = resolved.indexByLabel['Rate']
    unitCol = resolved.indexByLabel['Unit']
  } catch {
    throw new Error('Could not find S.No / Qty / Rate / Unit columns in the estimate.')
  }
  // Description isn't reliably labelled (and is often a merged cell with no
  // header of its own) — it's always the cell right after the serial number.
  const descCol = snoCol + 1
  if (descCol >= header.length) {
    throw new Error('Could not find a description column next to the S.No column in the estimate.')
  }

  const items: EstimateWorkItem[] = []
  let block: { row: string[]; gridRow: number }[] = []

  function resolveBlock() {
    if (block.length === 0) return
    const leadRow = block[0]
    const description = norm(leadRow.row[descCol])
    const measureRows = block.filter(({ row }) => {
      return norm(row[rateCol]) !== '' && norm(row[unitCol]) !== '' && norm(row[qtyCol]) !== ''
    })
    // A block with more than one measurement row is a single item broken
    // into several depth/size variants — e.g. "Drilling of tube wells…"
    // followed by "0 to 30 mtrs", "30 to 60 mtrs", "60 to 90 mtrs" rows,
    // each with its own Qty/Rate but sharing the lead row's description.
    // Emit one BOQ item per variant (tagging each with its own row label)
    // instead of collapsing them all into just the last measurement row.
    const multipleVariants = measureRows.length > 1
    for (const { row, gridRow } of measureRows) {
      const rate = norm(row[rateCol])
      const unit = norm(row[unitCol])
      const quantity = norm(row[qtyCol])
      const label = norm(row[descCol])
      const itemDescription = multipleVariants && label && label !== description ? `${description}( ${label})` : description
      if (itemDescription && quantity) {
        items.push({
          description: itemDescription,
          quantity,
          rate,
          unit,
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
    }
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

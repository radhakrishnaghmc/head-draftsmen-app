import { resolveColumns } from './columnMatch'
import type { ColumnSpec, ColumnEmbeddings } from './columnMatch'
import type { EstimateWorkItem } from './estimateExtract'
import { correctHeaderTypos } from './estimateExtract'

// A dedicated extractor for FLAT electrical estimates — a plain BOQ where each
// item is a single row (Sl.No | Description | Qty | Rate | Unit | Amount), with
// no No.s/L/B/D dimension sub-rows the civil "detailed & abstract" estimates
// use. Kept separate from core/estimateExtract.ts's block algorithm so an
// electrical sheet is read row-for-row (simpler, and robust to the wide variety
// of one-line electrical rate items) rather than pattern-matched for summary
// lines that a flat sheet doesn't have.

function norm(s: unknown): string {
  return String(s ?? '')
    .replace(/[\r\n]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function isNumeric(s: string): boolean {
  const t = s.replace(/,/g, '').trim()
  return t !== '' && Number.isFinite(Number(t))
}

// Only Qty and Rate are required to identify the item table; the rest help but
// are optional (an electrical sheet may not label a Unit column, and its
// Description header is often mis-spelled — "Descriptiion" — so it's matched
// loosely and can fall back to the column right of the serial number).
const ELECTRICAL_COLUMN_SPECS: ColumnSpec[] = [
  { label: 'Quantity', patterns: [/qty|quantity/i] },
  { label: 'Rate', patterns: [/rate/i] },
  { label: 'Description', patterns: [/descrip|particular|item\b/i], optional: true },
  { label: 'Serial Number', patterns: [/^sl?\.?\s*no\.?$|serial/i], optional: true },
  { label: 'Unit', patterns: [/unit|uom|^per$/i], optional: true },
  { label: 'Amount', patterns: [/amount/i], optional: true }
]

// The row that begins the estimate's cost build-up (Sub Total / GST / Total /
// Add … / Seigniorage / LS …) — where the item list ends. Matched against the
// row's description OR quantity cell, since these labels land in either column.
const ABSTRACT_RE =
  /^(sub\s*[-–—]?\s*total|grand\s*total|total\b|add\b|.*\bgst\b|labour\s*cess|\bnac\b|seign[io]rage|smft|dmft|permit|tpqc|\bls\b|unfor[e]?seen|provision|say\b)/i

/**
 * Read the line items out of a flat electrical estimate's grid. Resolves the
 * Qty/Rate columns (required) plus Description/Serial/Unit/Amount (best-effort),
 * then walks the rows after the header: a row with a numeric Quantity AND a
 * numeric Rate is an item; a row whose Description/Qty cell reads a cost-build-up
 * label (Sub Total, GST, Total, Add …) ends the list. Section-header rows (a
 * description but no Qty/Rate) are skipped. Throws when it can't find the
 * Qty/Rate columns at all (the sheet isn't a recognisable rate table).
 */
export function extractElectricalEstimateItems(
  grid: string[][],
  headerRowIndex: number,
  embeddings?: ColumnEmbeddings
): EstimateWorkItem[] {
  const header = (grid[headerRowIndex] ?? []).map((c) => correctHeaderTypos(norm(c)))
  let resolved
  try {
    resolved = resolveColumns(header, ELECTRICAL_COLUMN_SPECS, embeddings)
  } catch {
    throw new Error('Could not find the Qty and Rate columns in the electrical estimate.')
  }
  const qtyCol = resolved.indexByLabel['Quantity']
  const rateCol = resolved.indexByLabel['Rate']
  const unitCol = resolved.indexByLabel['Unit']
  const amountCol = resolved.indexByLabel['Amount']
  const serialCol = resolved.indexByLabel['Serial Number']
  const descCol =
    resolved.indexByLabel['Description'] ?? (serialCol != null ? serialCol + 1 : Math.max(0, qtyCol - 1))

  const items: EstimateWorkItem[] = []
  for (let r = headerRowIndex + 1; r < grid.length; r++) {
    const row = grid[r] ?? []
    const desc = norm(row[descCol])
    const qty = norm(row[qtyCol])
    const rate = norm(row[rateCol])
    const amount = amountCol !== undefined ? norm(row[amountCol]) : ''

    // Reached the cost build-up (label may sit in the description or qty cell).
    if ((desc && ABSTRACT_RE.test(desc)) || (qty && ABSTRACT_RE.test(qty))) break

    // A real item row carries a numeric Quantity and a numeric Rate (abstract
    // rows have an Amount but no Qty/Rate, so they never qualify).
    if (desc && isNumeric(qty) && isNumeric(rate)) {
      items.push({
        description: desc,
        quantity: qty,
        rate,
        unit: unitCol !== undefined ? norm(row[unitCol]) : '',
        ...(amount ? { estimateAmount: amount } : {})
      })
    }
  }
  return items
}

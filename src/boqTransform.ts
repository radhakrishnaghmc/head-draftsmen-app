import type { ExcelTable } from '@core/types'
import { type EstimateWorkItem, UNIT_RE } from '@core/estimateExtract'
import { BOQ_HEADERS } from '@core/boqHeaders'
import { resolveColumns } from '@core/columnMatch'
import type { ColumnSpec, ColumnEmbeddings } from '@core/columnMatch'

// Tries each pattern in order (most specific first) and returns the first
// header any of them match — lets a column be found by its exact expected
// name, falling back to looser aliases real-world BOQs commonly use instead.
// Used only against the app's own fixed BOQ_HEADERS (never a messy uploaded
// file), so no embedding fallback is needed here — see
// BOQ_SCHEDULE_A_COLUMN_SPECS below for the uploaded-BOQ case, which does.
function findColumn(headers: string[], patterns: RegExp | RegExp[], label: string): string {
  const { indexByLabel } = resolveColumns(headers, [{ label, patterns: Array.isArray(patterns) ? patterns : [patterns] }])
  return headers[indexByLabel[label]]
}

// The columns Schedule A needs from an *uploaded* BOQ — unlike the app's own
// BOQ_HEADERS, this file could have been hand-built by anyone with any
// column naming, so these are resolved together (see boqToScheduleA) with an
// embedding fallback available for whatever no pattern here anticipated.
export const BOQ_SCHEDULE_A_COLUMN_SPECS: ColumnSpec[] = [
  { label: 'Estimate Quantity', patterns: [/quantity/i, /\bqty\b/i] },
  {
    label: 'Item Detailed Specification Description',
    patterns: [
      /detailed\s+specification\s+description/i,
      /item\s+description/i,
      /description\s+of\s+item/i,
      /\bparticulars\b/i,
      /\bdescription\b/i
    ]
  },
  { label: 'Rate', patterns: [/\brate\b/i] },
  // Detailed CMC/departmental estimates (No.s/L/B/D layout) often carry no
  // Unit/UOM header at all — the unit (Cum/Sqm/…) sits in an unlabelled column,
  // filled only on each item's summary line. So UOM is optional here and, when
  // no header matches, detected from the data (see detectUnitColumn).
  { label: 'UOM', patterns: [/\buom\b/i, /\bunit\b/i, /\bunits\b/i], optional: true },
  { label: 'Amount', patterns: [/\bamount\b/i] }
]

/**
 * When an uploaded BOQ/estimate has no labelled Unit/UOM column, find the
 * unlabelled column that most often carries a recognised unit token
 * (Cum/Sqm/Rmt/…) across the real item rows — tie-broken toward the column just
 * after Rate, where the unit usually sits. Returns the header name, or undefined
 * when no column looks like a unit (Schedule A then leaves Units blank rather
 * than failing).
 */
function detectUnitColumn(
  headers: string[],
  itemRows: Record<string, string>[],
  claimed: Set<string>,
  rateCol: string
): string | undefined {
  const rateIdx = headers.indexOf(rateCol)
  let bestCol: string | undefined
  let bestHits = 0
  let bestDist = Infinity
  headers.forEach((col, idx) => {
    if (claimed.has(col)) return
    let hits = 0
    for (const row of itemRows) {
      const cell = (row[col] ?? '').trim()
      if (cell && UNIT_RE.test(cell)) hits++
    }
    const dist = Math.abs(idx - (rateIdx + 1))
    if (hits > bestHits || (hits === bestHits && hits > 0 && dist < bestDist)) {
      bestHits = hits
      bestCol = col
      bestDist = dist
    }
  })
  return bestHits > 0 ? bestCol : undefined
}

/**
 * Which of BOQ_SCHEDULE_A_COLUMN_SPECS' columns (if any) only resolved via
 * the embedding fallback rather than a regex — worth a "please double-check
 * this match" notice in the UI. Returns [] both when nothing needed the
 * fallback and when column resolution failed outright (boqToScheduleA
 * itself is the source of truth for that failure).
 */
export function boqColumnsMatchedViaEmbedding(headers: string[], embeddings?: ColumnEmbeddings): string[] {
  try {
    return resolveColumns(headers, BOQ_SCHEDULE_A_COLUMN_SPECS, embeddings).viaEmbedding
  } catch {
    return []
  }
}

const SCHEDULE_A_HEADERS = [
  'Item No.',
  'Probable Quantity',
  'Description of item',
  'Estimated Rate in Rs. & Ps.',
  'Units',
  'Estimated Amount in Rs'
]

function isNumericLike(v: string): boolean {
  const s = v.trim().replace(/,/g, '')
  return s !== '' && Number.isFinite(Number(s))
}

// A repeated column-header row: estimates/BOQs that span pages often reprint
// their "Qty | Description of work | Rate | Per" header mid-sheet, and that row
// would otherwise slip through as a bogus line item (a "Description of work"
// row whose amount comes out as #VALUE!). The description cell carrying a bare
// header label is the tell — a genuine line item never has one.
function isHeaderLabel(v: string): boolean {
  return /^(description(\s+of\s+(work|items?|specification))?|particulars|item(\s+no\.?)?|sl?\.?\s*no\.?)$/i.test(v.trim())
}

/**
 * Build Schedule A from an uploaded BOQ: Item No. is a running serial number,
 * and the remaining columns are pulled straight from the matching BOQ column
 * (quantity, detailed specification description, rate, UOM, amount). Columns
 * are matched by regex first; when `embeddings` is supplied and a header
 * matches none of them, the closest-matching header by semantic similarity
 * is used instead (e.g. a hand-built BOQ with unusual column names) — see
 * core/columnMatch.ts.
 */
export function boqToScheduleA(boq: ExcelTable, embeddings?: ColumnEmbeddings): ExcelTable {
  let indexByLabel: Record<string, number>
  try {
    indexByLabel = resolveColumns(boq.headers, BOQ_SCHEDULE_A_COLUMN_SPECS, embeddings).indexByLabel
  } catch (e) {
    throw new Error(`${e instanceof Error ? e.message : String(e)} in the uploaded BOQ.`)
  }
  const qtyCol = boq.headers[indexByLabel['Estimate Quantity']]
  const descCol = boq.headers[indexByLabel['Item Detailed Specification Description']]
  let rateCol = boq.headers[indexByLabel['Rate']]
  const amountCol = boq.headers[indexByLabel['Amount']]

  // BOQ exports often carry a trailing subtotal row (blank description, just
  // a lone amount) or a stray note row (a description but no figures) below
  // the real items — a genuine line item always has both a description and
  // at least a quantity or an amount, so require both to filter those out.
  const itemRows = boq.rows.filter((row) => {
    const desc = (row[descCol] ?? '').trim()
    if (desc === '') return false
    if (isHeaderLabel(desc)) return false // a reprinted column-header row, not a line item
    return isNumericLike(row[qtyCol] ?? '') || isNumericLike(row[amountCol] ?? '')
  })

  // UOM is optional: use its labelled column when present, otherwise detect the
  // unlabelled unit column from the data (detailed estimates leave it unheaded).
  let uomCol: string | undefined =
    indexByLabel['UOM'] != null ? boq.headers[indexByLabel['UOM']] : undefined
  if (uomCol == null) {
    const claimed = new Set(Object.values(indexByLabel).map((i) => boq.headers[i]))
    uomCol = detectUnitColumn(boq.headers, itemRows, claimed, rateCol)
  }

  // Some source files have Rate and UOM filled in under the wrong header
  // (Rate holding "Cum"/"Rmt" text, UOM holding the numeric rate) — if the
  // "Rate" column is mostly non-numeric while "UOM" is mostly numeric across
  // the actual item rows, they've been swapped, so swap them back.
  if (uomCol && itemRows.length > 0) {
    const uom = uomCol
    const rateNumeric = itemRows.filter((row) => isNumericLike(row[rateCol] ?? '')).length
    const uomNumeric = itemRows.filter((row) => isNumericLike(row[uom] ?? '')).length
    if (uomNumeric > rateNumeric) {
      ;[rateCol, uomCol] = [uom, rateCol]
    }
  }

  const rows = itemRows.map((row, i) => ({
    'Item No.': String(i + 1),
    'Probable Quantity': row[qtyCol] ?? '',
    'Description of item': row[descCol] ?? '',
    'Estimated Rate in Rs. & Ps.': row[rateCol] ?? '',
    Units: (uomCol ? row[uomCol] : '') ?? '',
    'Estimated Amount in Rs': row[amountCol] ?? ''
  }))

  return { id: `schedule-a-${Date.now()}`, name: 'Schedule A', path: '', headers: SCHEDULE_A_HEADERS, rows }
}

/**
 * Build a filled BOQ (standard column layout) from the work items extracted
 * from an estimate:
 *  - Estimate Quantity / Rate / UOM ← copied straight from the estimate.
 *  - Item Detailed Specification Description ← the estimate's description.
 *  - Work Type ← the description's first 3 words.
 *  - Item Short Description ← the description's first 50 characters.
 *  - APSS / Morth Cl. Number ← "NA" for every row.
 *  - Amount ← Estimate Quantity × Rate.
 */
export function buildBoqFromEstimate(
  items: EstimateWorkItem[],
  boqHeaders: string[] = BOQ_HEADERS
): ExcelTable {
  const qtyCol = findColumn(boqHeaders, /estimate\s*quantity/i, 'Estimate Quantity')
  const descCol = findColumn(
    boqHeaders,
    /detailed\s*specification\s*description|item\s*detailed/i,
    'Item Detailed Specification Description'
  )
  const workTypeCol = findColumn(boqHeaders, /work\s*type/i, 'Work Type')
  const shortDescCol = findColumn(boqHeaders, /item\s*short\s*description/i, 'Item Short Description')
  const apssCol = findColumn(boqHeaders, /apss|morth/i, 'APSS / Morth Cl. Number')
  const rateCol = findColumn(boqHeaders, /\brate\b/i, 'Rate')
  const uomCol = findColumn(boqHeaders, /\buom\b/i, 'UOM')
  const amountCol = findColumn(boqHeaders, /\bamount\b/i, 'Amount')

  const rows = items.map((it) => {
    const description = it.description.trim()
    const qty = Number(String(it.quantity).replace(/,/g, '')) || 0
    const rate = Number(String(it.rate).replace(/,/g, '')) || 0
    const row: Record<string, string> = {}
    for (const h of boqHeaders) row[h] = ''
    row[qtyCol] = it.quantity
    row[descCol] = description
    row[workTypeCol] = description.split(/\s+/).slice(0, 3).join(' ')
    row[shortDescCol] = description.slice(0, 50)
    row[apssCol] = 'NA'
    row[rateCol] = it.rate
    row[uomCol] = it.unit
    row[amountCol] = (qty * rate).toFixed(2)
    return row
  })

  return { id: `boq-${Date.now()}`, name: 'BOQ', path: '', headers: boqHeaders, rows }
}

/**
 * The BOQ's own Amount total: each item's Quantity × Rate rounded to the
 * nearest rupee, then summed — matching the live formula the BOQ template
 * uses per row (ROUND(rate*qty,0)) before its own SUM total. This is the
 * same figure the estimate's own item amounts sum to, i.e. the work's ECV.
 */
export function computeEcvFromItems(items: EstimateWorkItem[]): number {
  return items.reduce((sum, it) => {
    const qty = Number(String(it.quantity).replace(/,/g, '')) || 0
    const rate = Number(String(it.rate).replace(/,/g, '')) || 0
    return sum + Math.round(qty * rate)
  }, 0)
}

const WORK_NAME_ROW_RE = /^name of work:\s*(.+)$/i

/**
 * Recover the work name a generated BOQ carries in its own file (written two
 * rows below the total by core/boqTemplate.ts) when that BOQ is later
 * re-uploaded to build a Schedule A — scans every cell of every row for the
 * "Name of Work: ..." marker, regardless of which column it landed in once
 * re-parsed (it never has a quantity/amount, so boqToScheduleA's own item
 * filter already excludes it from the Schedule A item table).
 */
export function extractWorkNameFromBoq(table: ExcelTable): string | undefined {
  for (const row of table.rows) {
    for (const header of table.headers) {
      const m = WORK_NAME_ROW_RE.exec((row[header] ?? '').trim())
      if (m) return m[1].trim()
    }
  }
  return undefined
}

/**
 * Parse a flat BOQ (a Bill of Quantities uploaded directly, not re-derived from
 * an estimate) into the same EstimateWorkItem shape the estimate extractor
 * produces — so a BOQ can drive the very same BOQ / Schedule A / Material
 * outputs. Reuses boqToScheduleA's column resolution and item-row filtering, so
 * it accepts the same range of BOQ layouts. Returns [] when no line items
 * resolve (the caller then treats the file as neither an estimate nor a BOQ).
 */
export function boqToItems(boq: ExcelTable, embeddings?: ColumnEmbeddings): EstimateWorkItem[] {
  const sched = boqToScheduleA(boq, embeddings)
  return sched.rows.map((r) => ({
    description: r['Description of item'] ?? '',
    quantity: r['Probable Quantity'] ?? '',
    rate: r['Estimated Rate in Rs. & Ps.'] ?? '',
    unit: r['Units'] ?? '',
    estimateAmount: r['Estimated Amount in Rs'] ?? ''
  }))
}

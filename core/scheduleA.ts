import * as XLSX from 'xlsx'
import type { ExcelTable } from './types'
import { computeWorkAmounts, indianDigitGroups, tenderPercentFromRow } from './worksAmounts'
import { rankByEmbedding } from './embeddingMatch'

// A match below this score is treated as "no real match" — same threshold
// used for placeholder/column matching elsewhere (core/columnMatch.ts).
const EMBEDDING_THRESHOLD = 0.5

export interface ScheduleAItem {
  itemNo: string
  quantity: string
  description: string
  rate: string
  units: string
  amount: string
}

/** Work-detail fields pulled from the Works List row for the selected Win Code. */
export interface ScheduleAMeta {
  nameOfWork?: string
  estimateAmount?: string
  ecvAmount?: string
  contractAmount?: string
  contractorName?: string
  tenderPercentage?: string
  /** The Executive Engineer's circle office, e.g. "Nizampet Circle-58" — replaces the template's sample circle in the signature block. */
  circle?: string
  /** The zone, e.g. "Quthbullapur" — replaces the template's sample zone in the preamble. */
  zone?: string
  /** Corporation abbreviation (e.g. "CMC") — fills the SE Schedule A's {{Corp}} token in its Superintending-Engineer signature block. */
  corporation?: string
}

/**
 * Build Schedule-A meta fields from a Works List row. "Name of the Agency"
 * and "Name of the Contractor" are the same concept in this app — the Works
 * List only ever has an Agency column, and it's what fills Schedule-A's
 * "Name of the Contractor" field.
 *
 * Estimate/ECV/Contract Amount are run through the app's Lakhs-to-rupees +
 * Indian-numbering money rules (core/worksAmounts.ts) rather than passed
 * through as the raw Lakhs figure the Works List stores them as — Schedule
 * A's own labels already print "Rs.", so these come back as just the
 * Indian-grouped figure + "/-" (e.g. "45,00,000/-"), not a second "Rs".
 * Contract Amount is the computed ECV-net-of-tender-percentage figure, not
 * whatever raw value happens to be in that Works List cell — and is left
 * unset (blank in the output) when Tender Percentage isn't available yet,
 * rather than assuming 0%. ECV itself is left unset the same way when the
 * Works List row's ECV cell is blank — never substituted with the estimate.
 */
export function metaFromWorksRow(row: Record<string, string>): ScheduleAMeta {
  const c = computeWorkAmounts(row)
  // The circle office line the way the documents print it: "<Circle> Circle-<CNO>"
  // (e.g. "Nizampet Circle-58"), so Schedule A's signature isn't left stamped
  // with the template's sample circle.
  const circleName = (row['Circle'] ?? '').trim()
  const cno = (row['Circle number'] ?? row['CNO'] ?? '').trim()
  const circle = circleName ? (cno ? `${circleName} Circle-${cno}` : circleName) : undefined
  return {
    nameOfWork: row['Name of the work'],
    estimateAmount: `${indianDigitGroups(c.estimate)}/-`,
    ecvAmount: c.ecv !== null ? `${indianDigitGroups(c.ecv)}/-` : undefined,
    contractAmount: c.contractAmount !== null ? `${indianDigitGroups(c.contractAmount)}/-` : undefined,
    contractorName: row['Name of the Agency'],
    tenderPercentage: tenderPercentFromRow(row) || undefined,
    circle,
    zone: (row['Zone'] ?? '').trim() || undefined
  }
}

export interface WorksRowMatch {
  row: Record<string, string>
  /** True when the match came from semantic similarity (embeddings) rather than an exact name match — worth a "please verify" flag, same as column-matching elsewhere. */
  matchedViaAi: boolean
}

/**
 * Find the Works List row whose "Name of the work" matches (case- and
 * whitespace-insensitive) `workName` — used to auto-fill Schedule-A meta from
 * a work name detected in an uploaded BOQ, instead of requiring a manually
 * picked Win Code.
 *
 * A BOQ's title-block work name can differ slightly in wording from the
 * Works List's own entry (abbreviations, punctuation, rephrasing) — when the
 * exact match fails and `embeddings` are supplied (workName + every
 * candidate row name, both embedded by the caller), the closest semantic
 * match above the threshold is used instead.
 */
export function findWorksRowByName(
  worksTable: ExcelTable,
  workName: string,
  embeddings?: { workNameVector: number[]; rowNameVectors: number[][] }
): WorksRowMatch | undefined {
  const nameHeader = worksTable.headers.find((h) => h.trim().toLowerCase() === 'name of the work')
  if (!nameHeader) return undefined
  const norm = (s: string) => s.trim().toLowerCase().replace(/\s+/g, ' ')
  const target = norm(workName)
  if (!target) return undefined

  const idx = worksTable.rows.findIndex((r) => norm(r[nameHeader] ?? '') === target)
  if (idx !== -1) return { row: worksTable.rows[idx], matchedViaAi: false }

  if (embeddings) {
    const [best] = rankByEmbedding(embeddings.workNameVector, embeddings.rowNameVectors)
    if (best && best.score >= EMBEDDING_THRESHOLD) {
      return { row: worksTable.rows[best.index], matchedViaAi: true }
    }
  }
  return undefined
}

/** Pull the mapped Schedule-A item rows back out of the preview table built by boqToScheduleA. */
export function rowsToScheduleAItems(table: ExcelTable): ScheduleAItem[] {
  return table.rows.map((row) => ({
    itemNo: row['Item No.'] ?? '',
    quantity: row['Probable Quantity'] ?? '',
    description: row['Description of item'] ?? '',
    rate: row['Estimated Rate in Rs. & Ps.'] ?? '',
    units: row['Units'] ?? '',
    amount: row['Estimated Amount in Rs'] ?? ''
  }))
}

function numOrStr(s: string): number | string {
  const n = Number(String(s).replace(/[,%]/g, '').trim())
  return Number.isFinite(n) && s.trim() !== '' ? n : s
}

/**
 * Build the full Schedule-A row layout — title, preamble, and metadata rows
 * around the item table — matching the standard tender Schedule-A layout.
 * Name of work, estimate/ECV/contract amount, contractor, and tender-quoted %
 * are filled from the Works List when available (meta); the signature block
 * is left blank since we have no source for it. Used for both the on-screen
 * preview and the saved workbook, so they always match.
 */
export function buildScheduleARows(
  items: ScheduleAItem[],
  meta?: ScheduleAMeta,
  signatoryTitle: string = 'Executive Engineer'
): (string | number)[][] {
  const total = items.reduce((sum, it) => sum + (Number(String(it.amount).replace(/,/g, '')) || 0), 0)
  // Fallback only when there's no Works List match at all — the item table's
  // own total, Indian-formatted the same way metaFromWorksRow's figures are.
  // Never used to paper over a matched row's genuinely blank ECV — that
  // stays blank rather than being shown as if it were a real figure.
  const totalDisplay = meta ? '' : `${indianDigitGroups(total)}/-`

  return [
    ['SCHEDULE-A'],
    ['BILL OF QUANTITIES'],
    ['Preamble'],
    [
      '1. The Bill of Quantities shall be read in conjunction with the Instructions to Tenderers, General Conditions of Contract and Conditions of Particular Application, Technical Specifications and Specification Drawings.'
    ],
    [
      '2. It is to be expressly understood that the measured work is to be taken net, notwithstanding any custom or practice, to the Specification Drawings or as may be ordered from time to time by the Executive Engineer, and the cost calculated by measurement.'
    ],
    [
      "3. The premium/discount tendered in Part I of the Bill of Quantities shall, except in so far as is otherwise provided for under the Contract, cover all construction plant (Contractor's equipment during construction), materials, erection, maintenance and all other things necessary for the proper execution of the work."
    ],
    [
      '4. The whole cost of complying with the provisions of the Contract shall be included in the items provided in the priced Bill of Quantities, and where no items are provided the cost shall be deemed to be distributed among the items entered in the Bill of Quantities.'
    ],
    ['5. General directions and descriptions of work and materials are not necessarily repeated nor summarised in the Bill of Quantities.'],
    ['6. This schedule is liable to alterations by omissions, deductions, or additions at the discretion of the department.'],
    ['7. Errors will be corrected as follows: where there is a discrepancy between figures and words, the words will govern.'],
    ['Name of work:', meta?.nameOfWork ?? ''],
    ['Estimate Amount: Rs.', '', meta?.estimateAmount ?? totalDisplay],
    ['ECV Amount: Rs.', '', meta?.ecvAmount ?? totalDisplay],
    ['Contract Amount: Rs.', '', meta?.contractAmount ?? ''],
    ['Name of the Contractor :', meta?.contractorName ?? ''],
    [
      'Item No.',
      'Probable Quantity',
      'Description of item',
      'Estimated Rate in Rs. & Ps.',
      'Units',
      'Estimated Amount in  Rs'
    ],
    [1, 2, 3, 4, 5, 6],
    ...items.map((it) => [
      numOrStr(it.itemNo),
      numOrStr(it.quantity),
      it.description,
      numOrStr(it.rate),
      it.units,
      numOrStr(it.amount)
    ]),
    ['', '', '', '', '', total],
    ['', '', 'Tender Quoted', numOrStr(meta?.tenderPercentage ?? ''), meta?.tenderPercentage ? '%' : '', 'Less'],
    ['', '', 'Excess:'],
    ['', '', 'Less:'],
    ['', '', '', signatoryTitle]
  ]
}

/** Serialize the Schedule-A rows to an .xlsx workbook buffer for saving to disk. */
export function buildScheduleAWorkbook(items: ScheduleAItem[], meta?: ScheduleAMeta): Buffer {
  const aoa = buildScheduleARows(items, meta)
  const sheet = XLSX.utils.aoa_to_sheet(aoa)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, 'Schedule A')
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer
}

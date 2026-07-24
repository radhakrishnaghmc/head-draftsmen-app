import { extractLabeledField } from './estimateExtract'
import type { EstimateWorkItem } from './estimateExtract'

const ESTIMATE_AMOUNT_LABEL_RE = /^(?:estimate\s*amount|amount\s*of\s*estimate)\b\s*:?\s*(.*)$/i
const ECV_LABEL_RE = /^(?:ecv|estimate\s*contract\s*value)\b\s*:?\s*(.*)$/i

/**
 * The work's overall estimate amount, in Lakhs (matching how every other
 * amount field in this app is entered) — read from a labeled "Estimate
 * Amount" (or, failing that, "ECV") cell in the estimate's own title block
 * when present, otherwise computed by summing the item amounts (Qty x Rate)
 * and converting that rupee total back to Lakhs.
 */
export function extractEstimateAmountLakhs(
  grid: string[][],
  headerRowIndex: number,
  items: EstimateWorkItem[]
): number {
  const labeled =
    extractLabeledField(grid, headerRowIndex, ESTIMATE_AMOUNT_LABEL_RE) ??
    extractLabeledField(grid, headerRowIndex, ECV_LABEL_RE)
  if (labeled) {
    const n = Number(labeled.replace(/,/g, '').trim())
    if (Number.isFinite(n)) return n
  }

  const totalRupees = items.reduce((sum, it) => {
    const qty = Number(String(it.quantity).replace(/,/g, '')) || 0
    const rate = Number(String(it.rate).replace(/,/g, '')) || 0
    return sum + qty * rate
  }, 0)
  return Math.round((totalRupees / 100000) * 100) / 100
}

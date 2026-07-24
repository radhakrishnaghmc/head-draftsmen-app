import { listParagraphs, setParagraphText } from './docx-edit'
import { lakhsToRupees, indianDigitGroups, formatRupees } from './worksAmounts'

// Placeholder tokens present in the bundled Bid Document template, and the
// zip parts they can appear in — {{Circle}} repeats in the running footer
// (footer2/footer5), everything else lives in the body (document.xml).
const DOC_PARTS = ['word/document.xml', 'word/footer2.xml', 'word/footer5.xml']

export interface BidDocumentWorkItem {
  /** Serial number within the tender notice's item table (1-based) — becomes "BID Document N". */
  serial: number
  name: string
  /** Estimate amount as entered on the Works List, in Lakhs (e.g. "45" = Rs 45,00,000). */
  amount: string
  /** Estimate Amount ECV, in Lakhs — EMD @ 1% is computed from this when present, otherwise from `amount`. */
  ecv?: string
  zone?: string
  circle?: string
  completionPeriod?: string
}

export interface BidDocumentInput {
  nitNo: string
  /** DD.MM.YYYY */
  dated: string
  /** DD.MM.YYYY, e.g. "16.07.2026" — time (2:00 P.M) is appended automatically. */
  downloadStartDate: string
  /** DD.MM.YYYY — time (2:00 P.M) is appended automatically. */
  downloadEndDate: string
  work: BidDocumentWorkItem
}

export { lakhsToRupees }

/** Replace every occurrence of literal `{{Token}}` placeholders across the given zip parts, preserving run formatting (via core/docx-edit's diff-based rewrite) — mirrors the substitution approach in core/tenderNotice.ts, generalized to headers/footers. */
function fillPlaceholders(buffer: Buffer, swaps: Array<[string, string]>): Buffer {
  let current = buffer
  for (const part of DOC_PARTS) {
    let paragraphs: string[]
    try {
      paragraphs = listParagraphs(current, part)
    } catch {
      continue // template has no such part (e.g. fewer footers)
    }
    for (const [token, value] of swaps) {
      if (token === value) continue
      let list = paragraphs
      for (let i = 0; i < list.length; i++) {
        if (!list[i].includes(token)) continue
        const next = list[i].split(token).join(value)
        current = setParagraphText(current, i, next, list[i], part)
        list = listParagraphs(current, part)
      }
      paragraphs = list
    }
  }
  return current
}

/**
 * Fill the bundled Bid Document template for a single work: the NIT-level
 * fields (NIT No., Dated, download window) plus that work's own Name,
 * Estimate Amount, ECV, Zone, Circle, Completion Period and EMD @ 1%
 * (computed from the Estimate Amount ECV when the Works List row has one,
 * per the app's Lakhs-to-rupees/Indian-numbering money rules — see
 * core/worksAmounts.ts).
 *
 * The template's cover page has no "Rs:" label of its own before
 * {{Estimate Amount}}, so that occurrence gets the full "Rs 1,00,000/-".
 * The NIT body's {{ECV}} and {{EMD 1%}} sit right after a "Rs:" the
 * template already prints, so those get just the Indian-grouped figure +
 * "/-", not a second "Rs".
 */
export function fillBidDocument(buffer: Buffer, input: BidDocumentInput): Buffer {
  const estimateRupees = lakhsToRupees(input.work.amount)
  const ecvRupees = input.work.ecv?.trim() ? lakhsToRupees(input.work.ecv) : estimateRupees
  const emdRupees = Math.round(ecvRupees * 0.01)

  const swaps: Array<[string, string]> = [
    ['{{Name of the work}}', input.work.name],
    ['{{Estimate Amount}}', formatRupees(estimateRupees)],
    ['{{ECV}}', `${indianDigitGroups(ecvRupees)}/-`],
    ['{{Completion period}}', input.work.completionPeriod ?? ''],
    ['{{EMD 1%}}', `${indianDigitGroups(emdRupees)}/-`],
    ['{{ Download Start Date}}', `${input.downloadStartDate} at 2:00 P.M`],
    ['{{ Download End Date}}', `${input.downloadEndDate} at 2:00 P.M`],
    ['{{ Price Bid Opening}}', `${input.downloadEndDate} at 2:30 P.M`],
    ['{{Dated}}', input.dated],
    ['{{Nit no.}}', input.nitNo],
    ['{{Zone}}', input.work.zone ?? ''],
    ['{{Circle}}', input.work.circle ?? '']
  ]

  return fillPlaceholders(buffer, swaps)
}

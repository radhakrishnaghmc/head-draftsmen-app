import { listParagraphs, setParagraphText } from './docx-edit'
import { lakhsToRupees, rupeesFromCell, indianDigitGroups, formatRupees } from './worksAmounts'
import { zoneAbbr } from './loaSe'
export { stripItemNoTag } from './tenderAgents/shared'

// Placeholder tokens present in the bundled Bid Document template, and the
// zip parts they can appear in — {{Circle}} repeats in the running footer
// (footer2/footer5), everything else lives in the body (document.xml).
const DOC_PARTS = ['word/document.xml', 'word/footer2.xml', 'word/footer5.xml']

// Same idea for the SE-office (zone-only, no Circle) template — {{Zone}}
// repeats in its running footer (footer1/footer2/footer4; footer3/footer5
// carry no placeholders in the bundled template).
const SE_DOC_PARTS = ['word/document.xml', 'word/footer1.xml', 'word/footer2.xml', 'word/footer4.xml']

export interface BidDocumentWorkItem {
  /** Serial number within the tender notice's item table (1-based) — becomes "BID Document N". */
  serial: number
  name: string
  /** Estimate amount as entered on the Works List, in Lakhs (e.g. "45" = Rs 45,00,000). */
  amount: string
  /** ECV, in Lakhs — EMD @ 1% is computed from this. Left blank (never computed from `amount`) when ECV isn't available yet. */
  ecv?: string
  zone?: string
  circle?: string
  completionPeriod?: string
  /** SE template only — the NIT's item number, e.g. "3" for "(Item No.3)". Falls back to a "(Item No.N)" tag parsed out of `name`, then to `serial`, when not given (a Works List "Item No" column, if present). */
  itemNo?: string
  /** SE template only — full value as the office types it, e.g. "29/SE/QBZ/CMC/2026-27" (mirrors how `nitNo` is supplied whole). */
  tsNo?: string
  /** SE template only — DD-MM-YYYY, the Technical Sanction date. */
  tsDate?: string
  /** SE template only — which authority approved Administrative Sanction; the Bid Document's own AS line reads differently for each. Defaults to 'commissioner'. */
  asAuthority?: 'zonal' | 'commissioner'
  /** SE template only — DD.MM.YYYY (or DD-MM-YYYY), the Administrative Sanction date. Left out of the AS line entirely when blank, matching the office's own samples. */
  asDate?: string
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
function fillPlaceholders(buffer: Buffer, swaps: Array<[string, string]>, parts: string[] = DOC_PARTS): Buffer {
  let current = buffer
  for (const part of parts) {
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
 * (computed from the Works List row's ECV, per the app's Lakhs-to-rupees/
 * Indian-numbering money rules — see core/worksAmounts.ts). ECV and EMD @ 1%
 * are left blank — never computed from the estimate instead — when the
 * work's ECV isn't available yet, since the two are distinct figures.
 *
 * The template's cover page has no "Rs:" label of its own before
 * {{Estimate Amount}}, so that occurrence gets the full "Rs 1,00,000/-".
 * The NIT body's {{ECV}} and {{EMD 1%}} sit right after a "Rs:" the
 * template already prints, so those get just the Indian-grouped figure +
 * "/-", not a second "Rs".
 */
export function fillBidDocument(buffer: Buffer, input: BidDocumentInput): Buffer {
  const estimateRupees = lakhsToRupees(input.work.amount)
  // ECV is stored in rupees on the Works List (Amount of estimate is in Lakhs).
  const ecvRupees = input.work.ecv?.trim() ? rupeesFromCell(input.work.ecv) : null
  const emdRupees = ecvRupees !== null ? Math.round(ecvRupees * 0.01) : null

  const swaps: Array<[string, string]> = [
    ['{{Name of the work}}', input.work.name],
    ['{{Estimate Amount}}', formatRupees(estimateRupees)],
    ['{{ECV}}', ecvRupees !== null ? `${indianDigitGroups(ecvRupees)}/-` : ''],
    ['{{Completion period}}', input.work.completionPeriod ?? ''],
    ['{{EMD 1%}}', emdRupees !== null ? `${indianDigitGroups(emdRupees)}/-` : ''],
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

/** A Lakhs figure (as entered on the Works List, e.g. "63") to the 2-decimal display the SE documents use, e.g. "63.00". Blank/unparseable -> ''. */
function lakhsFixed2(lakhs: string | undefined): string {
  const n = Number((lakhs ?? '').replace(/,/g, '').trim())
  return Number.isFinite(n) ? n.toFixed(2) : ''
}

/**
 * ECV in Lakhs, TRUNCATED (not rounded) to 2 decimals — matches the SE
 * office's own convention in the source documents this template was built
 * from (e.g. an ECV of Rs 49,98,557 prints as "49.98 Lakhs", not the
 * arithmetically-rounded "49.99").
 */
function ecvLakhsTruncated(ecvRupees: number): string {
  return (Math.floor(ecvRupees / 1000) / 100).toFixed(2)
}

/** Pulls "3" out of a work name ending "...(Item No.3)" — the office's own convention for naming items within a multi-item NIT. Undefined when no such tag is present. */
export function extractItemNo(name: string): string | undefined {
  return /\(Item No\.?\s*(\d+)\)/i.exec(name)?.[1]
}

/**
 * The Administrative Sanction line, which reads structurally differently
 * per authority (not just a date swap) — the Zonal Commissioner form names
 * the zone, the plain Commissioner form doesn't. Either form omits its date
 * clause entirely when the date isn't available yet, matching how the
 * office's own documents look before the AS date is filled in.
 */
function adminSanctionLine(authority: 'zonal' | 'commissioner', zoneAbbrCode: string, asDate: string | undefined): string {
  const date = (asDate ?? '').trim()
  if (authority === 'zonal') {
    return `Administrative Sanction approved by the Zonal Commissioner, ${zoneAbbrCode}, CMC` + (date ? ` vide Dt. ${date}.` : '')
  }
  return `Administrative Sanction approved by the Commissioner, CMC` + (date ? ` Dt:${date}` : '')
}

/**
 * Fill the bundled SE-office (Superintending Engineer, zone-level — no
 * Circle) Bid Document template — used in place of `fillBidDocument` when
 * the issuing office has a Zone but no Circle of its own (see `seMode`
 * elsewhere in the app: `!!office.zone && !office.circle`). This document
 * carries several fields the EE template doesn't (Item No., Technical
 * Sanction No./Date, Administrative Sanction, and the Amount of
 * Estimate/ECV shown again in Lakhs), so — unlike the EE variant — it needs
 * `work.itemNo`/`tsNo`/`tsDate`/`asAuthority`/`asDate` filled in by the
 * caller; ECV and EMD @ 1% are still computed the same way, and still left
 * blank rather than falling back to the estimate, when ECV isn't available.
 */
export function fillSeBidDocument(buffer: Buffer, input: BidDocumentInput): Buffer {
  const ecvRupees = input.work.ecv?.trim() ? rupeesFromCell(input.work.ecv) : null
  const emdRupees = ecvRupees !== null ? Math.round(ecvRupees * 0.01) : null
  const zoneAbbrCode = zoneAbbr(input.work.zone)
  const itemNo = input.work.itemNo?.trim() || extractItemNo(input.work.name) || String(input.work.serial)

  const swaps: Array<[string, string]> = [
    ['{{Name of the work}}', input.work.name],
    ['{{Item No}}', itemNo],
    ['{{Amount of Estimate}}', lakhsFixed2(input.work.amount)],
    ['{{Administrative Sanction}}', adminSanctionLine(input.work.asAuthority ?? 'commissioner', zoneAbbrCode, input.work.asDate)],
    ['{{TS No}}', input.work.tsNo ?? ''],
    ['{{TS Date}}', input.work.tsDate ?? ''],
    ['{{ECV}}', ecvRupees !== null ? `${indianDigitGroups(ecvRupees)}/-` : ''],
    ['{{ECV Lakhs}}', ecvRupees !== null ? ecvLakhsTruncated(ecvRupees) : ''],
    ['{{EMD 1%}}', emdRupees !== null ? `${indianDigitGroups(emdRupees)}/-` : ''],
    ['{{Completion period}}', input.work.completionPeriod ?? ''],
    ['{{Dated}}', input.dated],
    ['{{Nit no.}}', input.nitNo],
    ['{{Zone}}', input.work.zone ?? ''],
    ['{{Zone Abbr}}', zoneAbbrCode]
  ]

  return fillPlaceholders(buffer, swaps, SE_DOC_PARTS)
}

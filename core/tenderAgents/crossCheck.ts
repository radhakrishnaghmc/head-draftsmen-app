import { detectTenderId } from './tenderId'
import { detectNitNoAndDate } from './nitNoAndDate'
import { detectNameOfWork } from './nameOfWork'
import { detectEcv } from './ecv'
import { detectTenderPercentage } from './tenderPercentage'
import { detectContractValue } from './contractValue'
import { detectBidSubmissionStartDate, detectBidSubmissionEndDate } from './bidSubmissionDates'
import { detectAgencyNames, detectAllBidders } from './agencyNames'
import { detectL1Agency } from './l1Agency'
import { extractEstimateFromTitle } from './estimateInTitle'
import type { NoteBidder } from '../noteSubmitted'

export interface TenderFields {
  tenderId?: string
  noticeNo?: string
  noticeDate?: string
  nameOfWork?: string
  ecvRupees?: number
  tenderPercentage?: number
  contractRupees?: number
  bidStart?: string
  bidClose?: string
  agencyNames: string[]
  allBidders: NoteBidder[]
  l1Agency?: string
  /** Non-blocking notes from cross-checking one agent's answer against
   * another's, or against a second, independent mention of the same fact
   * elsewhere on the page — never overrides a field's own value, only flags
   * it for a human to glance at. Empty when nothing looked off. */
  warnings: string[]
}

/**
 * NIT No and Tender ID are extracted by separate regexes that both scan the
 * same header text — if a future layout variant confuses one (e.g. cleanNit
 * failing to strip a wedged Tender ID the way it's had to be taught to for
 * several other real layouts already), the two won't naturally disagree the
 * way, say, ECV vs the work title do; they'll just silently carry the same
 * mistake. These are the two checks that can actually still fire given what
 * each agent already guarantees about its own output shape — a Tender ID is
 * always exactly 5-8 digits by construction (detectTenderId's own regex
 * enforces that before returning anything), so checking that again here
 * would be dead code; NIT No has no such guarantee (its raw-fallback path in
 * cleanNit can pass through un-validated text), so that one is worth it.
 */
function crossCheckIdentityShapes(tenderId: string | undefined, noticeNo: string | undefined, warnings: string[]): void {
  // "/DB/" and "/SE/" cover every office seen so far except one
  // (Serilingampally/Ameenpur's "Engg-<n>/CMC/<office>/<year>" — see
  // nitNoAndDate.ts) — that shape is a real, correctly-extracted NIT No,
  // not a malformed one, so it's excluded here too rather than warning on
  // every single document from that office.
  if (noticeNo && !/\/DB\/|\/SE\/|^Engg-\d+\//i.test(noticeNo)) {
    warnings.push(`NIT No "${noticeNo}" doesn't contain the expected "/DB/" or "/SE/" segment — worth a manual check.`)
  }
  if (tenderId && noticeNo && noticeNo.includes(tenderId)) {
    warnings.push(`NIT No "${noticeNo}" appears to still contain the Tender ID "${tenderId}" — the two may not have been separated cleanly.`)
  }
}

/** How far apart two rupee figures can be before they're worth flagging as
 * disagreeing, rather than just two roundings of the same real number — a
 * work title's "(Est amt: 2.00 Lakhs)" is always a rough, hand-typed
 * figure, never the precise ECV, so exact agreement isn't the bar. */
const ECV_DISAGREEMENT_TOLERANCE = 0.15

function crossCheckEcvAgainstTitle(nameOfWork: string | undefined, ecvRupees: number | undefined, warnings: string[]): number | undefined {
  if (!nameOfWork) return ecvRupees
  const titleEstimate = extractEstimateFromTitle(nameOfWork)
  if (titleEstimate == null) return ecvRupees

  if (ecvRupees == null) {
    warnings.push(
      `ECV wasn't found on the page directly — using the estimate mentioned in the work title (₹${titleEstimate.toLocaleString('en-IN')}) instead. Please verify.`
    )
    return titleEstimate
  }

  const diff = Math.abs(ecvRupees - titleEstimate) / Math.max(ecvRupees, titleEstimate)
  if (diff > ECV_DISAGREEMENT_TOLERANCE) {
    warnings.push(
      `ECV (₹${ecvRupees.toLocaleString('en-IN')}) and the estimate mentioned in the work title (₹${titleEstimate.toLocaleString('en-IN')}) disagree by more than ${Math.round(ECV_DISAGREEMENT_TOLERANCE * 100)}% — worth a manual check.`
    )
  }
  return ecvRupees
}

/**
 * Runs every L1-sheet agent against the same page lines and cross-checks
 * their answers against each other where a real relationship exists between
 * two fields, rather than treating each as a fully isolated guess:
 *
 *  - Name of Work sometimes carries its own rough estimate in the title text
 *    ("…(Est amt: 2.00 Lakhs)…") — cross-checked against the ECV agent's
 *    answer (used as a fallback when ECV wasn't found at all, or flagged as
 *    a disagreement worth a manual look when both are present but differ).
 *  - Tender ID and NIT No are cross-checked against each other's expected
 *    shape, since a mix-up between the two (one agent's regex anchoring on
 *    the wrong nearby label) is the one class of error neither can catch by
 *    only ever looking at its own page text.
 *
 * Every field keeps coming from its own dedicated agent — this function
 * only adds `warnings` on top, never silently overrides a value except the
 * one explicit ECV fallback documented above.
 */
export function runTenderAgents(lines: string[]): TenderFields {
  const tenderId = detectTenderId(lines)
  const { noticeNo, noticeDate } = detectNitNoAndDate(lines)
  const nameOfWork = detectNameOfWork(lines)
  const rawEcv = detectEcv(lines)

  const warnings: string[] = []
  crossCheckIdentityShapes(tenderId, noticeNo, warnings)
  const ecvRupees = crossCheckEcvAgainstTitle(nameOfWork, rawEcv, warnings)

  return {
    tenderId,
    noticeNo,
    noticeDate,
    nameOfWork,
    ecvRupees,
    tenderPercentage: detectTenderPercentage(lines),
    contractRupees: detectContractValue(lines),
    bidStart: detectBidSubmissionStartDate(lines),
    bidClose: detectBidSubmissionEndDate(lines),
    agencyNames: detectAgencyNames(lines),
    allBidders: detectAllBidders(lines),
    l1Agency: detectL1Agency(lines),
    warnings
  }
}

/**
 * Whether two independently-parsed documents (e.g. a tender's L1 sheet and
 * its separately-uploaded Intimation letter) plausibly describe the SAME
 * tender — Tender ID and NIT No are this app's two identity keys, so any
 * caller merging fields from more than one document (see
 * WorkOrderAgreementTab.tsx / GiveIntimationTab.tsx's existing pdfEval +
 * notice merge) should check this first. Agrees if EITHER key both sides
 * have agrees; a side missing a key never blocks a match on the other
 * (a Responsiveness page has no Tender ID some layouts, an Intimation
 * letter never has one at all) — but two present, DIFFERING values on the
 * same key is always a mismatch, never ignored.
 */
export function sameTender(a: Pick<TenderFields, 'tenderId' | 'noticeNo'>, b: Pick<TenderFields, 'tenderId' | 'noticeNo'>): boolean {
  const idsDisagree = !!a.tenderId && !!b.tenderId && a.tenderId !== b.tenderId
  const noticesDisagree = !!a.noticeNo && !!b.noticeNo && a.noticeNo !== b.noticeNo
  if (idsDisagree || noticesDisagree) return false
  const idsAgree = !!a.tenderId && !!b.tenderId && a.tenderId === b.tenderId
  const noticesAgree = !!a.noticeNo && !!b.noticeNo && a.noticeNo === b.noticeNo
  return idsAgree || noticesAgree
}

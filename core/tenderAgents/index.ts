/**
 * Eleven independent detectors, each responsible for exactly one field of a
 * tender's L1 (Commercial Evaluation) sheet or Intimation letter — Tender ID,
 * NIT No + Date, Name of Work, ECV, Tender Percentage, Contract Value, Bid
 * Submission Start/End Date, Name of the Agencies, L1 Agency, and Address of
 * the Agency.
 *
 * Why split them up: every one of these fields has, at some point, needed a
 * fix for a government office's PDF export punctuating or wrapping its own
 * layout slightly differently than every sample seen before. When all of
 * that lived in one function (the old parseTenderEvaluation), a fix for one
 * field's edge case sat right next to seven others' — easy to reason about
 * wrong, and a shared helper's bug (see priceBidRow.ts's isNameContinuation)
 * could affect fields whose own logic never changed. Each agent below owns
 * exactly one field, reads the page lines itself, and can be fixed, tested,
 * and reasoned about without touching the other ten.
 *
 * Genuinely shared plumbing (priceBidRow.ts's L-1 row scan, shared.ts's
 * trivial helpers) stays factored out on purpose — see priceBidRow.ts's own
 * comment for why four agents share one scan instead of each re-implementing
 * it.
 *
 * All 11 read the same `lines: string[]` shape, whatever document it came
 * from — see inputAdapters.ts for turning an OCR'd photo or a spreadsheet's
 * rows into that same shape (a digital PDF's text already reduces to it via
 * pdf.js's own reconstruction, no adapter needed).
 */
export { detectTenderId } from './tenderId'
export { detectNitNoAndDate, type NitNoAndDate } from './nitNoAndDate'
export { detectNameOfWork } from './nameOfWork'
export { detectEcv } from './ecv'
export { detectTenderPercentage } from './tenderPercentage'
export { detectContractValue } from './contractValue'
export { detectBidSubmissionStartDate, detectBidSubmissionEndDate } from './bidSubmissionDates'
export { detectAgencyNames, detectAllBidders } from './agencyNames'
export { detectL1Agency } from './l1Agency'
export { detectAgencyAddress, type AgencyAddressBlock } from './agencyAddress'
export { linesFromOcr, linesFromExcelRows, type OcrLikeLine } from './inputAdapters'
export { extractEstimateFromTitle } from './estimateInTitle'
export { runTenderAgents, sameTender, type TenderFields } from './crossCheck'

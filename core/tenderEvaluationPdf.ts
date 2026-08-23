import type { NoteBidder } from './noteSubmitted'
import { joinLines } from './tenderAgents/shared'
import { detectTenderId } from './tenderAgents/tenderId'
import { detectNitNoAndDate } from './tenderAgents/nitNoAndDate'
import { detectNameOfWork } from './tenderAgents/nameOfWork'
import { detectEcv } from './tenderAgents/ecv'
import { detectTenderPercentage } from './tenderAgents/tenderPercentage'
import { detectContractValue } from './tenderAgents/contractValue'
import { detectBidSubmissionStartDate, detectBidSubmissionEndDate } from './tenderAgents/bidSubmissionDates'
import { detectL1Agency } from './tenderAgents/l1Agency'
import { detectAllBidders } from './tenderAgents/agencyNames'

export interface TenderEvaluation {
  /** "Name of Work" from the tender header — the key matched against the Works List. */
  nameOfWork?: string
  tenderId?: string
  /** NIT / Notice number (the trailing "item N Dated:…" tail is dropped). */
  noticeNo?: string
  /** The "Dated:"/"Dt:" date carried in the NIT No line (e.g. "15.07.2026"), if present. */
  noticeDate?: string
  /** The page footer's "Server Time: 02/07/2026 …" date (bottom-right of the L1 sheet) — when the sheet was generated; used as the Note Submitted's Intimation date. Normalised to dd.mm.yyyy. */
  serverDate?: string
  /** Estimated Contract Value, in rupees (the portal reports rupees, not Lakhs). */
  ecvRupees?: number
  /** The L-1 (lowest / selected) bidder's company name. */
  l1AgencyName?: string
  /** L-1's quoted percentage, signed: positive = below estimate ("Less"), negative = above ("Excess"). */
  tenderPercentage?: number
  /** L-1's quoted amount (the awarded contract value), in rupees. */
  contractRupees?: number
  /** "Bid Submission Start Date & Time" from the sheet (e.g. "25/07/2026 06:02 PM") — the bid-document downloading start. */
  bidStart?: string
  /** "Bid Submission Closing Date" from the sheet (e.g. "01/08/2026 04:00 PM") — the downloading end / last date for receipt of bids. */
  bidClose?: string
}

/**
 * Parses the Telangana e-procurement portal's tender-evaluation page (its
 * "Preliminary Responsiveness" / "Commercial Evaluation" screens, saved as
 * PDF) from the text lines pdf.js reconstructs (see src/pdfToText.ts). The
 * Commercial Evaluation page carries the full picture — Tender ID, NIT No,
 * ECV, Name of Work, and a price-bid table whose L-1 row gives the winning
 * agency, its quoted percentage, and the awarded contract value; the
 * Responsiveness page lacks the price table, so those L-1 fields come back
 * undefined there.
 *
 * This is a thin composer, not where the extraction logic actually lives:
 * each field below comes from its own independent detector under
 * ./tenderAgents/ — see that folder's index.ts for why they're split up one
 * field per file. Deliberately calls the individual detectors directly
 * rather than tenderAgents/crossCheck.ts's runTenderAgents(): that
 * composer's ECV falls back to a rough estimate mentioned in the work
 * title when the ECV field itself isn't found, which is the right call for
 * a human reviewing a warning but must never happen here — a blank ECV
 * stays blank everywhere in this app (EMD/ASD/Contract Amount), never
 * silently substituting the Amount of estimate.
 */
export function parseTenderEvaluation(lines: string[]): TenderEvaluation {
  const joined = joinLines(lines)
  const { noticeNo, noticeDate } = detectNitNoAndDate(lines)

  // The page footer (bottom-right) prints "Server Time: 02/07/2026 03:59:31 PM"
  // — when this L1 sheet was generated. Anchored on the "Server Time"/"Server
  // Date" label so it's never confused with the notice or bid-submission dates
  // elsewhere on the page. Normalised to dd.mm.yyyy to match the note's other
  // dates. Used as the Note Submitted's Intimation date. (Not its own agent:
  // this field is read once here rather than duplicated, since nothing else
  // needs it and it shares no logic with any other field.)
  const server = /Server\s*(?:Time|Date)[^0-9]{0,20}(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})/i.exec(joined)

  return {
    tenderId: detectTenderId(lines),
    noticeNo,
    noticeDate,
    serverDate: server ? `${server[1]}.${server[2]}.${server[3]}` : undefined,
    nameOfWork: detectNameOfWork(lines),
    ecvRupees: detectEcv(lines),
    l1AgencyName: detectL1Agency(lines),
    tenderPercentage: detectTenderPercentage(lines),
    contractRupees: detectContractValue(lines),
    bidStart: detectBidSubmissionStartDate(lines),
    bidClose: detectBidSubmissionEndDate(lines)
  }
}

/**
 * Every bidder row of the price-bid table (not just L-1), in the page's order —
 * for the Note Submitted comparison table. See ./tenderAgents/agencyNames.ts
 * (the "Name of the Agencies" agent) for the actual extraction logic.
 */
export function parseAllBidders(lines: string[]): NoteBidder[] {
  return detectAllBidders(lines)
}

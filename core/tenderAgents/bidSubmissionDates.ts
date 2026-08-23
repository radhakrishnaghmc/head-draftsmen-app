import { joinLines } from './shared'

interface BidSubmissionWindow {
  bidStart?: string
  bidClose?: string
}

// Bid Submission Start / Closing dates. On the sheet both labels ("Bid
// Submission Start", "Bid Submission Closing") sit on one line and both
// date-times ("25/07/2026 06:02 PM", "01/08/2026 04:00 PM") on the next, in
// that order — so capture the first two date-times after the "Start" label.
// Shared by both detectors below (start and close are read from one match,
// not two independent scans, since they're positionally paired on the page —
// there's no separate "Closing" anchor reliable enough to find bidClose on
// its own without risking it grabbing an unrelated date elsewhere).
function detectBidSubmissionWindow(lines: string[]): BidSubmissionWindow {
  const joined = joinLines(lines)
  const bid =
    /Bid\s*Submission\s*Start[\s\S]{0,140}?(\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d{1,2}:\d{2}\s*[AP]M)[\s\S]{0,80}?(\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d{1,2}:\d{2}\s*[AP]M)/i.exec(
      joined
    )
  if (!bid) return {}
  return {
    bidStart: bid[1].replace(/\s+/g, ' ').trim(),
    bidClose: bid[2].replace(/\s+/g, ' ').trim()
  }
}

/**
 * Agent: Bid Submission Start Date & Time — when bid downloading opens (e.g.
 * "25/07/2026 06:02 PM"), read from the L1 sheet's header.
 */
export function detectBidSubmissionStartDate(lines: string[]): string | undefined {
  return detectBidSubmissionWindow(lines).bidStart
}

/**
 * Agent: Bid Submission Closing Date — the downloading end / last date for
 * receipt of bids (e.g. "01/08/2026 04:00 PM"), read from the L1 sheet's
 * header, right after the Start Date & Time.
 */
export function detectBidSubmissionEndDate(lines: string[]): string | undefined {
  return detectBidSubmissionWindow(lines).bidClose
}

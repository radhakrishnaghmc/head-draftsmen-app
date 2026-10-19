import { parseTenderEvaluation, type TenderEvaluation } from '../tenderEvaluationPdf'
import { parseParticipatingBidders, type ParticipatingBidder } from '../viewBiddersPdf'

/** What the Evaluation Sheet agent reads off one "View Bidders" page. */
export interface EvaluationSheetReading {
  header: TenderEvaluation
  bidders: ParticipatingBidder[]
}

/**
 * Agent: Evaluation Sheet — reads one "View Bidders" (Evaluation Stages)
 * page, as text lines from pdf.js or from OCR of a printed/photographed copy,
 * for the Evaluation Sheet tool. It is the single entry point that tool uses:
 * the header (Tender ID, NIT No + Date, Name of Work, ECV, bid dates) comes
 * from the per-field agents via parseTenderEvaluation, and the participating
 * bidders from the Supplier List reader (../viewBiddersPdf.ts).
 *
 * Printed portal pages go through OCR, which drops or glues spaces ("TenderID",
 * "NameofWork", "18/10/202609:29AM"); the readers it composes are written to
 * tolerate that, so fix new OCR variants in the field's own agent.
 */
export function readEvaluationSheetPage(lines: string[]): EvaluationSheetReading {
  return { header: parseTenderEvaluation(lines), bidders: parseParticipatingBidders(lines) }
}

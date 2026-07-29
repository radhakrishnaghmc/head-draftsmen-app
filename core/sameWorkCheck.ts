import type { IntimationNotice } from './intimationNotice'
import type { TenderEvaluation } from './tenderEvaluationPdf'

/**
 * Cross-checks that the two files uploaded to build a work's Intimation /
 * Work Order / Agreement — the portal "View Intimation Notice" HTML and the
 * L-1 selection / evaluation PDF — actually belong to the *same* work, so a
 * letter is never assembled from one work's agency and another work's tender.
 *
 * The agency name is the primary check: the Online Intimation is addressed to
 * the winning (L-1) agency, so the L1 selection form's L-1 agency must appear
 * in the Intimation. A disagreeing agency is a definite mismatch and takes
 * priority — it's checked before the NIT No.
 *
 * The NIT / tender notice number (the notice's `nitNo`, the PDF's `noticeNo`)
 * is checked as well: the same agency can win more than one tender, so even
 * when the agencies agree a differing NIT No is still a mismatch. When there's
 * no agency to compare, the NIT check stands on its own. If neither identifier
 * is available on both sides the status is 'unknown' (the caller can't verify,
 * so it doesn't block). Reservation isn't part of this — a work being reserved
 * for a category has no bearing on whether the two files describe the same
 * work.
 */
export type SameWorkStatus = 'match' | 'mismatch' | 'unknown'

export interface SameWorkResult {
  status: SameWorkStatus
  /** Which field settled it — for a human-readable explanation. */
  by?: 'nit' | 'agency'
  noticeNit?: string
  pdfNit?: string
  noticeAgency?: string
  pdfAgency?: string
}

/** Normalize an identifier for comparison: tighten spacing around / and -, collapse spaces, lowercase. */
function normId(s: string | undefined): string {
  return (s ?? '')
    .replace(/\s*([/\-])\s*/g, '$1')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
}

/**
 * Normalize a party/agency name for comparison: drop a leading "M/s"/"Messrs"
 * honorific, strip punctuation, collapse spaces, lowercase. Single-letter
 * tokens are kept — they're often a firm's real initials (e.g. "M V S
 * Constructions").
 */
function normName(s: string | undefined): string {
  return (s ?? '')
    .replace(/^\s*(m\s*\/\s*s\.?|messrs\.?)\s+/i, '')
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
}

/**
 * Whether the L-1 agency name appears in the Online Intimation's agency name
 * — the Intimation is addressed to the winning agency, so the two name the
 * same firm, but often worded differently (M/s prefix, punctuation, extra
 * address words on one side). Treat as the same agency when the names are
 * equal, one contains the other, or nearly all of the shorter name's words
 * are present in the longer.
 */
function sameAgency(noticeName: string, l1Name: string): boolean {
  if (noticeName === l1Name) return true
  if (noticeName.includes(l1Name) || l1Name.includes(noticeName)) return true
  const noticeTokens = new Set(noticeName.split(' '))
  const l1Tokens = l1Name.split(' ')
  const shared = l1Tokens.filter((t) => noticeTokens.has(t)).length
  return l1Tokens.length > 0 && shared / l1Tokens.length >= 0.8
}

export function checkSameWork(notice: IntimationNotice, pdf: TenderEvaluation): SameWorkResult {
  const noticeNit = normId(notice.nitNo)
  const pdfNit = normId(pdf.noticeNo)
  const noticeAgency = normName(notice.agencyName)
  const pdfAgency = normName(pdf.l1AgencyName)

  const base = {
    noticeNit: notice.nitNo,
    pdfNit: pdf.noticeNo,
    noticeAgency: notice.agencyName,
    pdfAgency: pdf.l1AgencyName
  }

  const nitKnown = !!noticeNit && !!pdfNit
  const agencyKnown = !!noticeAgency && !!pdfAgency

  // Agency is checked first: the Online Intimation is addressed to the winning
  // (L-1) agency, so the L1 sheet's agency must be the one named in the
  // Intimation. A disagreeing agency is a definite mismatch, and takes priority
  // over the NIT No.
  if (agencyKnown && !sameAgency(noticeAgency, pdfAgency)) {
    return { ...base, status: 'mismatch', by: 'agency' }
  }

  // Then the NIT No — when present on both sides and different, also a mismatch
  // (same agency can win more than one tender, so the NIT still has to agree).
  if (nitKnown && noticeNit !== pdfNit) {
    return { ...base, status: 'mismatch', by: 'nit' }
  }

  if (agencyKnown) return { ...base, status: 'match', by: 'agency' }
  if (nitKnown) return { ...base, status: 'match', by: 'nit' }
  return { ...base, status: 'unknown' }
}

/** A one-line, human-readable explanation of a mismatch for a warning notice. */
export function sameWorkMismatchMessage(r: SameWorkResult): string {
  if (r.by === 'nit') {
    return `These two files are for different works — the Online Intimation is for NIT No "${r.noticeNit}" but the L1 selection form is for NIT No "${r.pdfNit}". Upload both files for the same work.`
  }
  return `The L1 selection form's L-1 agency "${r.pdfAgency}" is not the agency named in the Online Intimation ("${r.noticeAgency}"). Upload the Online Intimation and L1 selection form for the same work and agency.`
}

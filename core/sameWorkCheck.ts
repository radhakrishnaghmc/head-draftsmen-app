import type { IntimationNotice } from './intimationNotice'
import type { TenderEvaluation } from './tenderEvaluationPdf'

/**
 * Cross-checks that the two files uploaded to build a work's Intimation /
 * Work Order / Agreement — the portal "View Intimation Notice" HTML and the
 * L-1 selection / evaluation PDF — actually belong to the *same* work, so a
 * letter is never assembled from one work's agency and another work's tender.
 *
 * The HTML notice carries no "Name of Work", so the reliable shared
 * identifier is the NIT / tender notice number, which both files print (the
 * notice as `nitNo`, the PDF as `noticeNo`). When a NIT No is present on both
 * sides it is authoritative: equal -> same work, different -> a definite
 * mismatch the caller should block on. When it's missing on either side the
 * NIT can't decide, so the agency name is used as a softer corroboration; and
 * if neither identifier is available on both sides the status is 'unknown'
 * (the caller can't verify, so it doesn't block). Reservation isn't part of
 * this — a work being reserved for a category has no bearing on whether the
 * two files describe the same work.
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

/** Normalize a party/agency name for comparison: collapse spaces, drop punctuation, lowercase. */
function normName(s: string | undefined): string {
  return (s ?? '')
    .replace(/[.,]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
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

  // NIT No is the authoritative key — present on both sides, it decides.
  if (noticeNit && pdfNit) {
    return { ...base, status: noticeNit === pdfNit ? 'match' : 'mismatch', by: 'nit' }
  }

  // No NIT No to compare — fall back to the agency name when both files name one.
  if (noticeAgency && pdfAgency) {
    return { ...base, status: noticeAgency === pdfAgency ? 'match' : 'mismatch', by: 'agency' }
  }

  return { ...base, status: 'unknown' }
}

/** A one-line, human-readable explanation of a mismatch for a warning notice. */
export function sameWorkMismatchMessage(r: SameWorkResult): string {
  if (r.by === 'nit') {
    return `These two files are for different works — the Online Intimation is for NIT No "${r.noticeNit}" but the L1 selection form is for NIT No "${r.pdfNit}". Upload both files for the same work.`
  }
  return `These two files look like they're for different works — the Online Intimation names agency "${r.noticeAgency}" but the L1 selection form's L-1 agency is "${r.pdfAgency}". Upload both files for the same work.`
}

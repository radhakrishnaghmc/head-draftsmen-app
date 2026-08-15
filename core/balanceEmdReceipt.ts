export interface BalanceEmdReceipt {
  /** The CURE portal's own receipt number. */
  receiptNo?: string
  /** e-Proc Tender ID (numeric, matches the L-1 sheet's own Tender ID). */
  tenderId?: string
  /** NIT / e-Proc tender notice number, e.g. "14/SE/QBZ/CMC/2026-27". */
  noticeNo?: string
  /** The NIT's own date, from the "…, Dated.17.07.2026" tail beside the notice number. */
  noticeDate?: string
  zone?: string
  circle?: string
  nameOfWork?: string
  agencyName?: string
  ecvRupees?: number
  /** Quoted percentage, negative when "Less" (matches TenderEvaluation's sign convention). */
  tenderPercentage?: number
  /** The 1.5% (or 2.5%, reserved works) balance EMD amount actually paid. */
  balanceEmdRupees?: number
  asdRupees?: number
  totalRupees?: number
  paymentId?: string
  /** When the payment was made, dd-mm-yyyy as printed on the receipt. */
  paymentDate?: string
  /** "SUCCESS" on a completed payment. */
  status?: string
}

function toNumber(s: string): number | undefined {
  const n = Number(s.replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : undefined
}

/** First regex match across the line array, label-anchored so an unrelated number elsewhere on the receipt is never picked up instead. */
function field(lines: string[], re: RegExp): string | undefined {
  for (const line of lines) {
    const m = re.exec(line)
    if (m) return m[1].trim()
  }
  return undefined
}

/**
 * Parses the Telangana CURE portal's "Balance EMD payment Receipt" — the
 * confirmation printed after an L-1 agency pays the balance 1.5%/2.5% EMD
 * online. Works from either a digital PDF's reconstructed text lines (see
 * src/pdfToText.ts) or OCR'd lines off a phone photo of the same receipt
 * (see electron/ocr.ts via the ocrPhotosToLines IPC call) — offices often
 * only have a photo, not a saved PDF. Every field is best-effort; anything
 * not found is left undefined so the caller can flag an unreadable upload.
 *
 * The receipt is a plain "Label : value" table, one row per line, except
 * "Name of the work" whose value wraps across several rows with the label
 * itself landing mid-wrap (a tall-row quirk of the source PDF) — handled
 * separately below by position rather than by regex.
 */
export function parseBalanceEmdReceipt(lines: string[]): BalanceEmdReceipt {
  const result: BalanceEmdReceipt = {}

  result.receiptNo = field(lines, /Receipt\s*No\.?\s*:?\s*(\S+)/i)
  result.tenderId = field(lines, /Tender\s*ID\s*:?\s*(\d{4,10})/i)
  result.paymentId = field(lines, /Payment\s*ID\s*:?\s*(\S+)/i)
  result.status = field(lines, /Transaction\s*Status\s*:?\s*(\S+)/i)
  result.paymentDate = field(lines, /Date\s*of\s*Payment\s*:?\s*(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})/i)
  result.zone = field(lines, /Zone\s*Name\s*:?\s*(.+)/i)
  result.circle = field(lines, /Circle\s*Name\s*:?\s*(.+)/i)

  // Tender Notice No. + its own date, e.g.
  // "14/SE/QBZ/CMC/2026-27, Dated.17.07.2026".
  const notice = field(lines, /Tender\s*Notice\s*No\.?\s*:?\s*(.+)/i)
  if (notice) {
    const m = /^(.*?),?\s*Dated\.?\s*:?\s*(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})/i.exec(notice)
    result.noticeNo = (m ? m[1] : notice).trim()
    if (m) result.noticeDate = m[2]
  }

  const agency = field(lines, /Name\s*of\s*the\s*Agency\s*:?\s*(.+)/i)
  if (agency) result.agencyName = agency

  const ecv = field(lines, /Estimated\s*Contract\s*Value\s*:?\s*Rs\.?\s*([\d,]+(?:\.\d+)?)/i)
  if (ecv) result.ecvRupees = toNumber(ecv)

  const pct = field(lines, /Tender\s*Percentage\s*:?\s*(-?\d+(?:\.\d+)?)/i)
  if (pct) result.tenderPercentage = Number(pct)

  const emd = field(lines, /Balance\s*EMD\s*:?\s*Rs\.?\s*([\d,]*\.?\d*)/i)
  if (emd) result.balanceEmdRupees = toNumber(emd) ?? 0

  const asd = field(lines, /\bASD\s*:?\s*Rs\.?\s*([\d,]*\.?\d*)/i)
  if (asd !== undefined) result.asdRupees = toNumber(asd) ?? 0

  const total = field(lines, /Total\s*Amount\s*:?\s*Rs\.?\s*([\d,]*\.?\d*)/i)
  if (total) result.totalRupees = toNumber(total) ?? 0

  // Name of the work — everything between "Circle Name" and "Type of the
  // work", minus the "Name of the work" label line itself (it can land as
  // its own line partway through the wrapped value, or be absent if OCR
  // merged it into a neighbouring line).
  const circleIdx = lines.findIndex((l) => /Circle\s*Name/i.test(l))
  const typeIdx = lines.findIndex((l) => /Type\s*of\s*the\s*work/i.test(l))
  if (circleIdx >= 0 && typeIdx > circleIdx) {
    const work = lines
      .slice(circleIdx + 1, typeIdx)
      .map((l) => l.replace(/Name\s*of\s*the\s*work/i, '').trim())
      .filter((l) => l.length > 0)
      .join(' ')
      .replace(/^:\s*/, '')
      .replace(/\s+/g, ' ')
      .trim()
    if (work) result.nameOfWork = work
  }

  return result
}

export interface TenderEvaluation {
  /** "Name of Work" from the tender header — the key matched against the Works List. */
  nameOfWork?: string
  tenderId?: string
  /** NIT / Notice number (the trailing "item N Dated:…" tail is dropped). */
  noticeNo?: string
  /** The "Dated:"/"Dt:" date carried in the NIT No line (e.g. "15.07.2026"), if present. */
  noticeDate?: string
  /** Estimated Contract Value, in rupees (the portal reports rupees, not Lakhs). */
  ecvRupees?: number
  /** The L-1 (lowest / selected) bidder's company name. */
  l1AgencyName?: string
  /** L-1's quoted percentage, signed: positive = below estimate ("Less"), negative = above ("Excess"). */
  tenderPercentage?: number
  /** L-1's quoted amount (the awarded contract value), in rupees. */
  contractRupees?: number
}

function toNumber(s: string): number | undefined {
  const n = Number(s.replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : undefined
}

// Tightens OCR/layout spacing inside an identifier so a code split across the
// PDF's two-column layout ("… Circle- 58/CMC …") rejoins ("…Circle-58/CMC…").
function tightenCode(s: string): string {
  return s
    .replace(/\s*([/\-])\s*/g, '$1')
    .replace(/\s+/g, ' ')
    .trim()
}

/**
 * Parses the Telangana e-procurement portal's tender-evaluation page (its
 * "Preliminary Responsiveness" / "Commercial Evaluation" screens, saved as
 * PDF) from the text lines pdf.js reconstructs (see src/pdfToText.ts). The
 * Commercial Evaluation page carries the full picture — Tender ID, NIT No,
 * ECV, Name of Work, and a price-bid table whose L-1 row gives the winning
 * agency, its quoted percentage, and the awarded contract value; the
 * Responsiveness page lacks the price table, so those L-1 fields come back
 * undefined there. Every field is best-effort against that page's specific
 * layout; anything not found is left undefined for the caller to skip.
 */
export function parseTenderEvaluation(lines: string[]): TenderEvaluation {
  const result: TenderEvaluation = {}
  const joined = lines.join(' ').replace(/\s+/g, ' ').trim()

  const tenderId = /Tender ID\s+(\d+)/i.exec(joined)
  if (tenderId) result.tenderId = tenderId[1]

  // NIT No spans the header's left column and is interrupted by the right
  // column's "Tender ID …"/"Notice Number" labels in reading order — strip
  // those out, then take what's between "NIT No." and the "item"/"Dated" tail.
  let cleaned = joined
  if (result.tenderId) cleaned = cleaned.replace(new RegExp(`Tender ID\\s+${result.tenderId}`, 'i'), ' ')
  cleaned = cleaned.replace(/Notice Number/gi, ' ')
  const nit = /NIT No\.?\s*(.+?)\s*(?:item\b|Dated\b|Name of Work\b)/i.exec(cleaned)
  if (nit) {
    const value = tightenCode(nit[1])
    if (value) result.noticeNo = value
  }

  // The NIT line carries the notice's own date as "Dated:15.07.2026" /
  // "Dt: 15-07-2026" — split out as the Tender notice Date. "Dated"/"Dt" is
  // the reliable anchor (bid-submission/server dates on the page carry no
  // such label), so an unrelated date is never misread as the notice date.
  const date = /\b(?:Dated|Dt)\b\.?\s*:?\s*(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})/i.exec(joined)
  if (date) result.noticeDate = date[1]

  const work = /Name of Work\s+(.+?)\s+(?:Tender Category|Tender Type|Estimated Contract)\b/i.exec(joined)
  if (work) result.nameOfWork = work[1].replace(/\s+/g, ' ').trim()

  // Price-bid table rows: "<Company> <ECV> <Less|Excess> <pct> <amount> L-<rank>".
  // The L-1 row is the winning bid.
  for (const line of lines) {
    const m = /^(.+?)\s+([\d,]+\.\d{2})\s+(Less|Excess)\s+([\d.]+)\s+([\d,]+\.\d{2})\s+L-?\s*(\d+)/i.exec(line.trim())
    if (!m) continue
    const rank = Number(m[6])
    if (rank !== 1) continue
    result.l1AgencyName = m[1].replace(/\s+/g, ' ').trim()
    if (result.ecvRupees == null) result.ecvRupees = toNumber(m[2])
    const pct = toNumber(m[4])
    if (pct != null) result.tenderPercentage = /excess/i.test(m[3]) ? -pct : pct
    result.contractRupees = toNumber(m[5])
    break
  }

  // ECV fallback (e.g. the Responsiveness page, which has no price table):
  // the "Estimated Contract Value" figure sits on the "OPEN - NCB …" line.
  if (result.ecvRupees == null) {
    const ecv = /OPEN\s*-\s*NCB\s+([\d,]+\.\d{2})/i.exec(joined)
    if (ecv) result.ecvRupees = toNumber(ecv[1])
  }

  return result
}

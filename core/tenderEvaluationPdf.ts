import type { NoteBidder } from './noteSubmitted'

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

function toNumber(s: string): number | undefined {
  const n = Number(s.replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : undefined
}

/** The nearest non-empty trimmed line before index i, or '' if none. */
function prevNonEmptyLine(lines: string[], i: number): string {
  for (let j = i - 1; j >= 0; j--) {
    const t = lines[j].trim()
    if (t) return t
  }
  return ''
}

/**
 * Whether `line` looks like the tail of a wrapped bidder name (e.g.
 * "Contractor" continuing "Kummary Renuka Devi Civil") rather than a number
 * row or one of the page's control / footer rows — used to reassemble an L-1
 * company name split across lines by the price-bid cell's line wrap.
 */
function isNameContinuation(line: string): boolean {
  const t = line.trim()
  if (!t || t.length > 40 || /\d/.test(t)) return false
  return !/^(back|save|reject|icons|view documents|dashboard|company name|department|refund|information technology|https?:)/i.test(t)
}

// The field labels that follow "Name of Work" on the page — reaching one ends
// the work-name value. "Works Percentage" is the merged value of the next two
// fields (Tender Category = "Works", Tender Evaluation Type = "Percentage")
// that this page's layout drops onto one line right after the work name.
function isWorkNameBoundary(line: string): boolean {
  return /^(works\s+percentage|tender\s+category|tender\s+type|tender\s+evaluation\s+type|estimated\s+contract|price\s+bid|bid\s+submission)\b/i.test(
    line.trim()
  )
}

// A line that's one of the page's own field labels (not part of the work
// name) — so it's never picked up as the value line preceding "Name of Work".
function isFieldLabelLine(line: string): boolean {
  const t = line.trim()
  return t === '' || /^(notice number|tender id|enquiry|ifb|current tender|commercial evaluation|preliminary responsiveness|name of work)\b/i.test(t)
}

/**
 * Pull the "Name of Work" value out of the reconstructed page lines. Two
 * layouts occur: the label and value on one line ("Name of Work <value> …
 * Tender Category"), and — when the value is long — the value cell wrapping
 * *around* its label, so the label lands on its own line between the two value
 * lines:
 *
 *   Laying of CC Road … (Ward No 276, Pragathi nagar Nizampet   <- value part 1
 *   Name of Work                                                <- label alone
 *   Circle-58, Quthbullapur Zone CMC) (Reserved for SC)         <- value part 2
 *   Works Percentage                                            <- next field
 *
 * The old single-line regex captured only the part after the label, dropping
 * the first line — producing a fragment that then mis-matched a different
 * work. So: take any text after the label on its line, else the value line
 * just above it, then append the following lines until the next field label.
 */
function extractNameOfWork(lines: string[]): string | undefined {
  const li = lines.findIndex((l) => /name of work/i.test(l))
  if (li < 0) return undefined

  const parts: string[] = []
  const after = lines[li].replace(/^.*?name of work\s*/i, '').trim()
  if (after) {
    parts.push(after)
  } else {
    // Label alone: the value's first part wraps ABOVE the label. Collect EVERY
    // preceding value line (a long title routinely spans two or more), walking
    // up until the previous field label (e.g. "Notice Number") or a boundary —
    // not just the single line immediately above, which used to drop the title
    // and leave only its "…under Municipal General Funds…" tail.
    const above: string[] = []
    for (let j = li - 1; j >= 0; j--) {
      const t = (lines[j] ?? '').trim()
      if (!t) continue
      if (isFieldLabelLine(t) || isWorkNameBoundary(t)) break
      above.push(t)
    }
    above.reverse()
    parts.push(...above)
  }
  for (let j = li + 1; j < lines.length; j++) {
    const t = lines[j].trim()
    if (!t) continue
    if (isWorkNameBoundary(t)) break
    parts.push(t)
    if (parts.join(' ').length > 400) break
  }
  const name = parts.join(' ').replace(/\s+/g, ' ').trim()
  return name || undefined
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
 * Reassemble a NIT No that pdf.js interleaved with the header's right column.
 * These NITs read "<code>/DB/EE/<place> Circle-<circleNo>/[QBZ/]CMC/<year>", but
 * the Tender ID, an "Enquiry/IFB/Tender Notice" label fragment, or a stray digit
 * run routinely lands between the wrapped "…Circle-" prefix and the
 * "<circleNo>/[QBZ/]CMC/<year>" tail (e.g. "…Circle- Enquiry/IFB/Tender 699588
 * 57/QBZ/CMC/2026-27"). Keep just the prefix and that tail, dropping whatever's
 * wedged between — and, because the tail ends at the year, this also trims any
 * trailing "date:…" / work-name text the value regex over-captured. Falls back
 * to the tightened raw value when the string doesn't fit this canonical shape.
 */
function cleanNit(raw: string): string {
  const m = /^(.*?Circle-)\s*(?:.*?\s)?(\d{1,3}\/(?:QBZ\/)?CMC\/\d{4}\s*-\s*\d{2,4})/i.exec(raw)
  return tightenCode(m ? `${m[1]}${m[2]}` : raw)
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

  // Tender ID sits in the header's right column; pdf.js emits it either as
  // "Tender ID <id>" or — when the value cell rounds a hair above its label —
  // as "<id> Tender ID" (value before label). Accept both orders.
  const tenderId = /Tender ID\s+(\d+)|(\d+)\s+Tender ID/i.exec(joined)
  if (tenderId) result.tenderId = tenderId[1] ?? tenderId[2]

  // NIT No spans the header's left column and is interrupted by the right
  // column ("Tender ID <id>" / the "Notice Number" label) landing between its
  // wrapped halves in reading order — e.g. "…Circle- 699549 Tender ID
  // 57/QBZ/…". Strip that Tender ID label+value in either order and the
  // "Notice Number" label so the NIT's two halves rejoin, then take what's
  // between "NIT No." and the "item"/"Dt"/"Dated" tail.
  const cleaned = joined
    .replace(/Tender ID\s+\d+|\d+\s+Tender ID/gi, ' ')
    .replace(/Notice Number/gi, ' ')
  // Capture through to the tail (or the line's end), then cleanNit drops any
  // Tender ID / "Enquiry/IFB/Tender" / date fragment wedged into the value.
  const nit = /NIT No\.?\s*(.+?)\s*(?:item\b|Dated\b|Dt\b|Name of Work\b|$)/i.exec(cleaned)
  if (nit) {
    const value = cleanNit(nit[1])
    if (value) result.noticeNo = value
  }

  // The NIT line carries the notice's own date as "Dated:15.07.2026" /
  // "Dt: 15-07-2026" — split out as the Tender notice Date. "Dated"/"Dt" is
  // the reliable anchor (bid-submission/server dates on the page carry no
  // such label), so an unrelated date is never misread as the notice date.
  const date = /\b(?:Dated|Dt)\b\.?\s*:?\s*(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})/i.exec(joined)
  if (date) result.noticeDate = date[1]

  // The page footer (bottom-right) prints "Server Time: 02/07/2026 03:59:31 PM"
  // — when this L1 sheet was generated. Anchored on the "Server Time"/"Server
  // Date" label so it's never confused with the notice or bid-submission dates
  // elsewhere on the page. Normalised to dd.mm.yyyy to match the note's other
  // dates. Used as the Note Submitted's Intimation date.
  const server = /Server\s*(?:Time|Date)[^0-9]{0,20}(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})/i.exec(joined)
  if (server) result.serverDate = `${server[1]}.${server[2]}.${server[3]}`

  // Bid Submission Start / Closing dates. On the sheet both labels ("Bid
  // Submission Start", "Bid Submission Closing") sit on one line and both
  // date-times ("25/07/2026 06:02 PM", "01/08/2026 04:00 PM") on the next, in
  // that order — so capture the first two date-times after the "Start" label.
  const bid = /Bid\s*Submission\s*Start[\s\S]{0,140}?(\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d{1,2}:\d{2}\s*[AP]M)[\s\S]{0,80}?(\d{1,2}\/\d{1,2}\/\d{2,4}\s+\d{1,2}:\d{2}\s*[AP]M)/i.exec(
    joined
  )
  if (bid) {
    result.bidStart = bid[1].replace(/\s+/g, ' ').trim()
    result.bidClose = bid[2].replace(/\s+/g, ' ').trim()
  }

  result.nameOfWork = extractNameOfWork(lines)
  if (!result.nameOfWork) {
    const work = /Name of Work\s+(.+?)\s+(?:Tender Category|Tender Type|Estimated Contract)\b/i.exec(joined)
    if (work) result.nameOfWork = work[1].replace(/\s+/g, ' ').trim()
  }

  // Price-bid table rows. The numeric core is
  // "<ECV> <Less|Excess> <pct> <amount> L-<rank>". The bidder's company name is
  // usually on the same line just before it, but a long name wraps onto its own
  // line(s), with the number cells landing on a line *between* the two name
  // lines (the numbers sit vertically centred against the wrapped name cell) —
  // e.g. "Kummary Renuka Devi Civil" / "3133583.00 Less 11.99 2757866.40 L-1" /
  // "Contractor". So when the number row carries no company prefix, take the
  // name from the line above, plus the line below when that's a name
  // continuation. The L-1 row is the winning bid.
  const priceCore = /([\d,]+\.\d{2})\s+(Less|Excess)\s+([\d.]+)\s+([\d,]+\.\d{2})\s+L-?\s*(\d+)\b/i
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim()
    const m = priceCore.exec(line)
    if (!m || Number(m[5]) !== 1) continue

    let company = line.slice(0, m.index).trim()
    if (!company) {
      const prev = prevNonEmptyLine(lines, i)
      const next = lines[i + 1]?.trim() ?? ''
      company = isNameContinuation(next) ? `${prev} ${next}`.trim() : prev
    }
    if (company) result.l1AgencyName = company.replace(/\s+/g, ' ').trim()
    if (result.ecvRupees == null) result.ecvRupees = toNumber(m[1])
    const pct = toNumber(m[3])
    if (pct != null) result.tenderPercentage = /excess/i.test(m[2]) ? -pct : pct
    result.contractRupees = toNumber(m[4])
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

/**
 * Every bidder row of the price-bid table (not just L-1), in the page's order —
 * for the Note Submitted comparison table. Uses the same row shape as
 * parseTenderEvaluation ("<Company> <ECV> <Less|Excess> <pct> <amount> L-<rank>")
 * but keeps all ranks. "Less" shows as "(-)pct" and "Excess" as a plain "pct",
 * matching how the office notes print the quoted percentage. Returns [] when the
 * page carries no price table (e.g. the Responsiveness screen).
 */
export function parseAllBidders(lines: string[]): NoteBidder[] {
  const out: NoteBidder[] = []
  for (const line of lines) {
    const m = /^(.+?)\s+([\d,]+\.\d{2})\s+(Less|Excess)\s+([\d.]+)\s+([\d,]+\.\d{2})\s+L-?\s*(\d+)/i.exec(line.trim())
    if (!m) continue
    const isExcess = /excess/i.test(m[3])
    out.push({
      sno: `${out.length + 1}.`,
      name: m[1].replace(/\s+/g, ' ').trim(),
      ecv: m[2].replace(/,/g, ''),
      pct: isExcess ? m[4] : `(-)${m[4]}`,
      tcv: m[5].replace(/,/g, ''),
      rank: `L-${m[6]}`
    })
  }
  return out
}

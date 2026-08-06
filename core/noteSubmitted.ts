// "Note Submitted" — the office file-noting sheet that carries one work through
// its six sequential notes (administrative sanction → technical sanction →
// tender invitation → tender opening + comparison → intimation → agreement),
// each closed by a right-tabbed "HD … EE" signature block. Built as HTML from a
// Works List row plus the bidder comparison (from the evaluation PDF), then
// converted to .docx for export via core/htmlToDocx.ts. See the plan in
// project memory (project_note_submitted_doc) for the field mapping.
import { computeWorkAmounts, tenderPercentFromRow } from './worksAmounts'

/** One row of the tender comparison table (display strings, echoed as-is). */
export interface NoteBidder {
  sno: string
  name: string
  ecv: string
  pct: string
  tcv: string
  rank: string
}

export interface NoteSubmittedData {
  /** Municipal body wording, e.g. "CMC" / "GHMC". */
  body: string
  circle: string
  workName: string
  /** "Amount of estimate" cell (Lakhs) — rendered as "Rs: 18.00 Lakhs". */
  estimateLakhs: string
  /** Administrative-sanction date (optional; the note tolerates a blank). */
  asDate: string
  financialYear: string
  /** The tender Notice No printed in note 3. The Tender Notice No and the NIT No
   * are the same reference, so `nitNo` mirrors this (the editor keeps them in
   * sync); likewise `nitDate` mirrors `tenderNoticeDate`. */
  tenderNoticeNo: string
  tenderNoticeDate: string
  /** Note 4's NIT No — same value as `tenderNoticeNo` (see above). */
  nitNo: string
  nitDate: string
  /** Newspapers the tender notice was published in (one sentence fragment). */
  newspapers: string
  /** Optional reason phrase for the rejected bidders (e.g. "low bid capacity"),
   * rendered inside note 4's "…(N) bidders rejected due to {reason}…" clause.
   * Free text — no count baked in; the count lives in `rejectedCount`. */
  qualificationNote: string
  /** Count of bidders made non-responsive / rejected (not shown in the price
   * table). Participants = bidders.length (qualified) + rejectedCount. */
  rejectedCount: number
  bidders: NoteBidder[]
  l1Name: string
  l1PctText: string
  l1Tcv: string
  intimationDate: string
  /** Reservation category on the work, e.g. "SC" / "ST" / "" — drives EMD exemption. */
  reservation: string
  /** L-1's tendered percentage as a signed number (positive = "Less"/below ECV). */
  l1PctNumber: number | null
  /** ECV in rupees — EMD (1.5%) and ASD ((%−25)%) are computed off this. */
  ecvRupees: number | null
  receiptNo: string
  receiptDate: string
}

const esc = (s: string): string =>
  String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')

/** Indian financial year for a date (1 April boundary): 2026-03-31 -> "2025-26". */
export function financialYearOf(d = new Date()): string {
  const y = d.getFullYear()
  const startY = d.getMonth() >= 3 ? y : y - 1
  return `${startY}-${String((startY + 1) % 100).padStart(2, '0')}`
}

/** EMD @ 1.5% is exempted for works reserved for a category (SC / ST). */
export function isReservedExempt(reservation: string): boolean {
  return /\b(SC|ST)\b/i.test(reservation)
}

/** A positioned line from the non-responsive sheet (see pdfToPositionedLines):
 * `x` is the line's leftmost item, `y` its top-to-bottom reading position. */
export interface NonRespLine {
  text: string
  x: number
  y: number
}

// The header row of the bidder table (its wrapped fragments) and the page's
// footer — used to bound the region that holds the actual bidder rows.
const NONRESP_HEADER = /company name.*comment/i
const NONRESP_FOOTER =
  /^(please select|back\b|save & continue|icons|view documents|dashboard \||department tender|refund emd|information technology|\*\* applicable|current tender)/i

/**
 * Count the bidders on a "List of Bidders Made Non-Responsive" sheet — and, for
 * a single bidder, its rejection reason — for note 4's rejected count.
 *
 * This reads the TABLE STRUCTURE, not the comment wording, because evaluators
 * routinely misspell the reason ("NOT RESPOSONVIE DUE TO LOW BID") — keyword
 * matching then miscounts (0, or double-counts a two-line comment). Instead:
 * the sheet has a company-name column on the far left and a Comments column on
 * the far right (the middle columns are checkboxes, so they carry no text).
 * Within the bidder region we take the right-hand (Comments) column — one text
 * block per bidder — and count blocks by their vertical gaps: small gaps are a
 * comment's own line-wraps, a big gap starts the next bidder's row.
 *
 * `detail` is filled only for a single bidder (several reasons don't compose
 * readably). Returns { count: 0, detail: '' } when no bidder rows are found.
 */
export function summarizeNonResponsiveness(lines: NonRespLine[]): { count: number; detail: string } {
  const clean = lines.filter((l) => l.text.trim())
  // Bound the region to the rows between the column header and the footer.
  const headerIdx = clean.findIndex((l) => NONRESP_HEADER.test(l.text))
  const startIdx = headerIdx >= 0 ? headerIdx + 1 : 0
  let endIdx = clean.length
  for (let i = startIdx; i < clean.length; i++) {
    if (NONRESP_FOOTER.test(clean[i].text)) {
      endIdx = i
      break
    }
  }
  const region = clean.slice(startIdx, endIdx)
  if (region.length === 0) return { count: 0, detail: '' }

  // The Comments column is the right-most; require a clear x-gap from the
  // name/middle columns so a sheet without one never yields phantom bidders.
  const maxX = Math.max(...region.map((l) => l.x))
  const minX = Math.min(...region.map((l) => l.x))
  if (maxX - minX < 150) return { count: 0, detail: '' }
  const comments = region.filter((l) => l.x >= maxX - 45)
  if (comments.length === 0) return { count: 0, detail: '' }

  // One comment block per bidder: cluster the (top-to-bottom) comment lines by
  // vertical gap. Threshold adapts to the line height (median gap) so it works
  // across font sizes, with a floor for the single-line-comment case.
  const gaps: number[] = []
  for (let i = 1; i < comments.length; i++) gaps.push(comments[i].y - comments[i - 1].y)
  const sorted = [...gaps].sort((a, b) => a - b)
  const median = sorted.length ? sorted[Math.floor(sorted.length / 2)] : 0
  const threshold = Math.max(18, median * 1.8)
  let count = 1
  for (const g of gaps) if (g > threshold) count++

  // Reason for a single bidder: the joined comment text, minus the leading
  // "(non|not) respo… [due to]" boilerplate (tolerant of the common misspelling).
  const detail =
    count === 1
      ? comments
          .map((c) => c.text)
          .join(' ')
          .replace(/\s+/g, ' ')
          .replace(/^(non|not)[\s-]*respo\S*\s*(due to|,|:)?\s*/i, '')
          .trim()
      : ''
  return { count, detail }
}

/** Plain whole-rupee string (no grouping), matching the sample notes' figures. */
const money = (n: number): string => String(Math.round(n))

/** "1.99", "4.99", "0.11" — the ASD percentage (tender% − 25), trimmed. */
const pct2 = (n: number): string => String(Math.round(n * 100) / 100)

/**
 * The variable clause in note 6, following the office's exact four shapes:
 *   reserved, ≤25%  → "EMD is Exempted"
 *   reserved, >25%  → "EMD is Exempted and submitted ASD amount of Rs.… of X%"
 *   open,     ≤25%  → "has submitted EMD amount of Rs.… of 1.5%"
 *   open,     >25%  → "…EMD … of 1.5% & ASD Amount of Rs.… of X%"
 * EMD = 1.5% × ECV, ASD = (tender% − 25) × ECV (matches computeWorkAmounts).
 */
function agreementClause(d: NoteSubmittedData): string {
  const pct = d.l1PctNumber
  const asd = d.ecvRupees != null && pct != null && pct > 25 ? Math.round(d.ecvRupees * ((pct - 25) / 100)) : 0
  const emd = d.ecvRupees != null ? Math.round(d.ecvRupees * 0.015) : null
  const asdPct = pct != null ? pct - 25 : null
  const emdAmt = emd != null ? money(emd) : '____________'
  const asdAmt = money(asd)
  const receipt = `vide Online Receipt No: ${esc(d.receiptNo) || RECEIPT_BLANK} Dt: ${esc(d.receiptDate) || '__________'}`

  if (isReservedExempt(d.reservation)) {
    const cat = esc(d.reservation.match(/\b(SC|ST)\b/i)?.[0]?.toUpperCase() ?? d.reservation)
    if (asd > 0 && asdPct != null) {
      return `the above said work is Reserved for ${cat}. Therefore, the EMD is Exempted and submitted ASD amount of Rs.${asdAmt}/- of ${pct2(asdPct)}% ${receipt}`
    }
    return `the above said work is Reserved for ${cat}. Therefore, the EMD is Exempted`
  }

  const base = `has submitted EMD amount of Rs.${emdAmt}/- of 1.5%`
  if (asd > 0 && asdPct != null) {
    return `${base} &amp; ASD Amount of Rs.${asdAmt}/- of ${pct2(asdPct)}% ${receipt}`
  }
  return `${base} ${receipt}`
}

// The right-tabbed bold "HD … EE" sign-off block that closes every note. Uses
// the HTML border="0" attribute (not CSS border:none) because html-to-docx
// ignores CSS border on tables and falls back to a default 2px border — the
// border="0" attribute is the only thing it honours to keep the block boxless.
const HD_EE =
  '<table border="0" style="width:100%;border-collapse:collapse;margin:2px 0 16px"><tr>' +
  '<td style="border:none;text-align:left;font-weight:bold">HD</td>' +
  '<td style="border:none;text-align:right;font-weight:bold">EE</td>' +
  '</tr></table>'

// Full-line hand-writable blanks: the newspapers clause (note 4) and the EMD
// Online Receipt No (note 6, a ~20-char number) each print a line-spanning
// underline to fill by hand when left empty, instead of a default / short gap.
const NEWSPAPER_BLANK = '_'.repeat(65)
const RECEIPT_BLANK = '_'.repeat(45)

// Note paragraphs are justified (text-align:justify → w:jc "both"); the Word
// 2007 sanitiser keeps <w:jc> in schema order after html-to-docx (see
// sanitizeHtmlDocxForWord2007).
const P = (html: string): string => `<p style="margin:0 0 6px;line-height:1.45;text-align:justify">${html}</p>`
const SUB = '<p style="margin:14px 0 4px;font-weight:bold">Submitted: -</p>'

/**
 * Builds the full Note Submitted body as HTML (no <html> wrapper) — safe to
 * inject into a preview element and to hand to convertHtmlToDocx for export.
 */
export function buildNoteSubmittedHtml(d: NoteSubmittedData): string {
  const b = esc(d.body) || 'CMC'
  const fy = esc(d.financialYear)
  const estimate = d.estimateLakhs.trim() ? `Rs: ${esc(d.estimateLakhs)} Lakhs` : 'Rs: __________ Lakhs'

  // Note 1 — administrative sanction
  const note1 = P(
    `It is to submit that the Commissioner/Zonal-Commissioner, ${b} has Accorded administrative sanction for the work of ` +
      `<b>${esc(d.workName)}</b> for Estimate amount of ${estimate} dt: ${d.asDate.trim() ? esc(d.asDate) : '_____________'} under general budget for the financial year ${fy}`
  )

  // Note 2 — technical sanction (fixed text)
  const note2 = P(
    `The items proposed in the estimate have been scrutinized with approved data current year ${fy}. ` +
      `As such the estimate is herewith put up for according technical sanction.`
  )

  // Note 3 — tender invitation
  const note3 = P(
    `The tender for the above said work is invited through this office vide tender Notice No ${esc(d.tenderNoticeNo)}, ` +
      `Dt: ${esc(d.tenderNoticeDate)} is Submitted for order please`
  )

  // Note 4 — tender opening + comparison table
  const rows = d.bidders
    .map(
      (r) =>
        '<tr>' +
        `<td style="border:1px solid #000;text-align:center">${esc(r.sno)}</td>` +
        `<td style="border:1px solid #000">${esc(r.name)}</td>` +
        `<td style="border:1px solid #000;text-align:center">${esc(r.ecv)}</td>` +
        `<td style="border:1px solid #000;text-align:center">${esc(r.pct)}</td>` +
        `<td style="border:1px solid #000;text-align:center">${esc(r.tcv)}</td>` +
        `<td style="border:1px solid #000;text-align:center">${esc(r.rank)}</td>` +
        '</tr>'
    )
    .join('')
  const table =
    '<table style="width:100%;border-collapse:collapse;margin:8px 0"><tr>' +
    ['S.No', 'Name of the Contractor', 'ECV', '% Quoted', 'TCV', 'Rank']
      .map((h) => `<td style="border:1px solid #000;text-align:center;font-weight:bold">${h}</td>`)
      .join('') +
    '</tr>' +
    rows +
    '</table>'
  // The price-bid table holds only the responsive (qualified) bidders, so the
  // number who *participated* is those plus the rejected/non-responsive ones.
  // Both counts are printed off the same numbers so they can never contradict
  // each other (the old code printed the qualified count as "participated").
  const qualified = d.bidders.length
  const rejected = Math.max(0, Math.trunc(Number(d.rejectedCount) || 0))
  const participants = qualified + rejected
  const reason = d.qualificationNote.trim().replace(/^[.\s]+/, '')
  // With rejections: "In that (N) bidder(s) rejected [due to …] and (M)
  // qualified as follows." Without: just "as follows." reading on into the
  // table below.
  let qualif: string
  if (rejected > 0) {
    const rNoun = rejected === 1 ? 'bidder' : 'bidders'
    const reasonClause = reason ? ` due to ${esc(reason)}` : ''
    qualif = `. In that (${rejected}) ${rNoun} rejected${reasonClause} and (${qualified}) qualified as follows.`
  } else {
    qualif = ' as follows.'
  }
  const note4 =
    P(
      `As per above note approved tenders have been called on e-procurement platform vide NIT No: ${esc(d.nitNo)}, ` +
        `Dt:${esc(d.nitDate)} &amp; e-procurement. tender notice has been published in ${d.newspapers.trim() ? esc(d.newspapers) : NEWSPAPER_BLANK}. ` +
        `In response to the tender notice (${participants}) bidders have participated${qualif}`
    ) +
    table +
    P(
      `The above One (1) tender, <b>${esc(d.l1Name)}</b> has quoted lowest rates on ECV i.e, ${esc(d.l1PctText)} % ` +
        `TCV is Rs.${esc(d.l1Tcv)}. Hence, submitted for approval.`
    )

  // Note 5 — intimation
  const note5 = P(
    `As per above note approved by EE ${esc(d.circle)}, Intimation letter is put up in favor <b>${esc(d.l1Name)}</b> for sign.`
  )

  // Note 6 — agreement (conditional EMD/ASD clause)
  const note6 =
    P(
      `In response to this office Intimation letter on DT: ${esc(d.intimationDate)}, <b>${esc(d.l1Name)}</b> ` +
        `${agreementClause(d)} and Rs.100/- Non-Judicial Stamp Paper for concluding agreement.`
    ) + P(`Hence, the draft agreement is herewith prepared and put up for approval.`)

  return (
    `<div style="font-family:'Bookman Old Style',Georgia,serif;font-size:12pt;line-height:1.5;color:#000">` +
    SUB +
    note1 +
    HD_EE +
    note2 +
    HD_EE +
    note3 +
    HD_EE +
    SUB +
    note4 +
    HD_EE +
    SUB +
    note5 +
    HD_EE +
    SUB +
    note6 +
    HD_EE +
    `</div>`
  )
}

/**
 * Pre-fills the note's data from a Works List row: work name, estimate, circle,
 * dates, NIT/notice numbers, L-1 agency/%/TCV, reservation and ECV — plus the
 * computed EMD/ASD inputs. The bidder table is seeded with the L-1 row alone;
 * callers replace it with the full list parsed from the evaluation PDF.
 * Fields absent from the row are left blank for the user to fill.
 */
export function noteSubmittedFromRow(row: Record<string, string>, defaultCircle = ''): NoteSubmittedData {
  const amounts = computeWorkAmounts(row)
  const reservation = (row['Reservation'] ?? '').trim() || (row['Name of the work'] ?? '').match(/reserved\s+for\s+([A-Za-z]+)/i)?.[1] || ''
  const pctRaw = tenderPercentFromRow(row).replace(/[%,\s]/g, '')
  const pctNum = pctRaw === '' ? null : Number(pctRaw)
  const l1Pct = pctNum == null || !Number.isFinite(pctNum) ? '' : pctNum > 0 ? `(-)${pctNum}` : String(pctNum)
  const ecv = amounts.ecv
  const tcv = amounts.contractAmount

  const l1Name = (row['Name of the Agency'] ?? '').trim()
  const bidders: NoteBidder[] = l1Name
    ? [
        {
          sno: '1.',
          name: l1Name,
          ecv: ecv != null ? ecv.toFixed(2) : '',
          pct: l1Pct,
          tcv: tcv != null ? tcv.toFixed(2) : '',
          rank: 'L-1'
        }
      ]
    : []

  return {
    body: 'CMC',
    circle: (row['Circle'] ?? '').trim() || defaultCircle,
    workName: (row['Name of the work'] ?? '').trim(),
    estimateLakhs: (row['Amount of estimate'] ?? '').trim(),
    asDate: '',
    financialYear: (row['Financial Year'] ?? '').trim() || financialYearOf(),
    tenderNoticeNo: (row['Tender Notice No'] ?? '').trim(),
    tenderNoticeDate: (row['Tender notice Date'] ?? '').trim(),
    nitNo: (row['Tender Notice No'] ?? '').trim(),
    nitDate: (row['Tender notice Date'] ?? '').trim(),
    // Left blank by default so the note prints a fill-in line — the office
    // writes the actual newspapers by hand (or types them in the editor).
    newspapers: '',
    qualificationNote: '',
    rejectedCount: 0,
    bidders,
    l1Name,
    l1PctText: l1Pct,
    l1Tcv: tcv != null ? tcv.toFixed(2) : '',
    intimationDate: (row['Intimation Date'] ?? '').trim(),
    reservation,
    l1PctNumber: pctNum != null && Number.isFinite(pctNum) ? pctNum : null,
    ecvRupees: ecv,
    receiptNo: '',
    receiptDate: ''
  }
}

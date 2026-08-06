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

/** A rejection-reason keyword — how each bidder's Comments cell opens. */
const REJECT_KEYWORD = /(non[- ]?responsive|not responsive|not satisfied|not uploaded|rejected)/i
/**
 * The portal's own static page chrome on the "List of Bidders Made
 * Non-Responsive" sheet that ALSO carries the rejection keyword and so must not
 * be counted as a bidder: the page title / section header ("List of Bidders
 * Made Non-Responsive [/Commercial Stage …]") and the instruction footer
 * ("Please Select Only such reasons for Non-Responsiveness …" / "Please Select
 * the reason for the Disqualification").
 */
const NONRESP_CHROME = /^(list of bidders|please select)/i

/**
 * Best-effort read of a "List of Bidders Made Non-Responsive" sheet — used to
 * pre-fill note 4's rejected count and reason (both fully editable afterwards).
 *
 * The sheet is a table: one row per rejected bidder, whose Comments cell opens
 * with a rejection keyword ("Non Responsive …"). pdf.js returns that cell as its
 * own line, with any wrapped continuation ("low Bid Capacity") on later lines
 * that carry no keyword — so counting keyword-bearing lines counts each bidder
 * once. The catch is the page's static chrome (title, section header, the
 * "Please Select … Non-Responsiveness" instruction) carries the keyword too;
 * those are excluded via NONRESP_CHROME, which is what made the old count
 * over-report (e.g. 5 for a 2-bidder sheet).
 *
 * `detail` is only filled for a single rejected bidder, where the reason cell's
 * lines compose into one clean phrase ("low Bid Capacity"); with several
 * bidders their reasons don't concatenate readably, so detail is left blank and
 * the note simply states the counts. Returns { count: 0, detail: '' } when no
 * rejected bidders are found.
 */
export function summarizeNonResponsiveness(lines: string[]): { count: number; detail: string } {
  const clean = lines.map((l) => l.trim()).filter(Boolean)
  const reasonStarts = clean.filter((l) => !NONRESP_CHROME.test(l) && REJECT_KEYWORD.test(l))
  const count = reasonStarts.length
  if (count !== 1) return { count, detail: '' }

  // Exactly one bidder: gather its Comments cell — every text-bearing line from
  // its reason start down to the instruction footer. Company-name fragments in
  // this region are ALL-CAPS (no lowercase), the reason fragments carry
  // lowercase, so keep only the lowercase-bearing lines, then drop the leading
  // "Non Responsive [due to]" boilerplate so it reads "…rejected due to {reason}".
  const startAt = clean.indexOf(reasonStarts[0])
  const fragments: string[] = []
  for (let i = startAt; i < clean.length; i++) {
    const l = clean[i]
    if (/^please select/i.test(l)) break
    if (/[a-z]/.test(l)) fragments.push(l)
  }
  const detail = fragments
    .join(' ')
    .replace(/\s+/g, ' ')
    .replace(/^non[- ]?responsive\s*(?:due to|,)?\s*/i, '')
    .trim()
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
  const receipt = `vide Online Receipt No: ${esc(d.receiptNo) || '____________________'} Dt: ${esc(d.receiptDate) || '__________'}`

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

const P = (html: string): string => `<p style="margin:0 0 6px;line-height:1.45">${html}</p>`
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
      `<b>${esc(d.workName)}</b> for Estimate amount of ${estimate} dt:${esc(d.asDate)} under general budget for the financial year ${fy}`
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
        `Dt:${esc(d.nitDate)} &amp; e-procurement. tender notice has been published in ${esc(d.newspapers)}. ` +
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
    newspapers: 'Andhra Jyothi daily Telugu Newspaper and the Pioneer Daily English Paper',
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

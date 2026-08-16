// Two SE-office (Zone, no Circle) notes issued alongside the Letter of
// Acceptance: the Bid Evaluation note (technical bid opened, bidders listed)
// and the Agency Approval note (financial bids compared, L-1 named). Both
// carry a bidder table whose row count varies per work, so — like the EE
// Note Submitted — they're built as HTML and converted via the app's generic
// convertHtmlToDocx (api.noteSubmittedDocx), not a fixed-row .docx template.
import type { NoteBidder } from './noteSubmitted'
import { indianDigitGroups } from './worksAmounts'

const esc = (s: string): string =>
  String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')

const INDENT = '&emsp;&emsp;&emsp;'
const P = (html: string): string => `<p style="margin:0 0 6px;line-height:1.45;text-align:justify">${INDENT}${html}</p>`

// html-to-docx ignores CSS border on <table>/<td> and falls back to a default
// 2px border unless the HTML border="0" attribute is set (see
// core/noteSubmitted.ts's HD_EE block) — used throughout below for the same
// reason.
const TABLE_STYLE = 'width:100%;border-collapse:collapse;margin:8px 0;table-layout:fixed'
const TH = (text: string, width?: string): string =>
  `<td style="border:1px solid #000;padding:4px 6px;font-weight:bold;text-align:center${width ? `;width:${width}` : ''}">${esc(text)}</td>`
const TD = (text: string, align: 'left' | 'center' | 'right' = 'left'): string =>
  `<td style="border:1px solid #000;padding:4px 6px;text-align:${align}">${text}</td>`

// `spacerLines` blank lines above the AE/Dy.EE/SE row — room for the three
// wet signatures on the printed copy, above and beyond the table's own margin.
const SIGN_OFF = (zoneAbbr: string, spacerLines = 0): string =>
  '<p>&nbsp;</p>'.repeat(spacerLines) +
  `<table border="0" style="width:100%;border-collapse:collapse;margin:24px 0 0"><tr>` +
  `<td style="border:none;text-align:left;font-weight:bold">AE</td>` +
  `<td style="border:none;text-align:center;font-weight:bold">Dy.EE</td>` +
  `<td style="border:none;text-align:right;font-weight:bold">SE (M), ${esc(zoneAbbr)}</td>` +
  `</tr></table>`

/** Fields shared by both notes' Sub/Ref block. */
export interface SeEvaluationRefs {
  zoneAbbr: string
  workName: string
  estimateLakhs: string
  /** "Zonal Commissioner, QBZ, CMC" or "Commissioner, CMC" — no trailing date. */
  asAuthority: string
  /** The Ref-1 sentence tail: asAuthority, plus ", dated:DD.MM.YYYY" only for the Commissioner form. */
  asRefLine: string
  tsNo: string
  tsDate: string
  financialYear: string
  nitNo: string
  nitDate: string
  itemNo: string
  techBidOpenDate: string
}

export interface BidEvaluationData extends SeEvaluationRefs {
  bidders: NoteBidder[]
  /** From the uploaded "List of Bidders Made Non-Responsive" sheet (see summarizeNonResponsiveness in core/noteSubmitted.ts). 0/undefined = not uploaded, or none found. */
  nonRespCount?: number
}

// "Ref:" plus every numbered point hang off the SAME left indent (the
// label's own width) — the label sits on item 1)'s line, not as a lone
// paragraph of its own with the numbered points starting fresh below it.
const REF_INDENT = '36px'
/** "Ref: 1)" (first item) or a bare "    2)" (every item after) — matching the office's own Ref-block numbering, hanging-indented so wrapped lines still align under the number. */
const refItem = (n: string, html: string, first = false): string =>
  `<p style="margin:0 0 2px;padding-left:${REF_INDENT};text-indent:-${REF_INDENT}">${first ? '<b>Ref:</b>&emsp;' : ''}${esc(n)}&emsp;${html}</p>`

function refBlockHtml(d: SeEvaluationRefs): string {
  return (
    refItem('1)', `Administrative sanction approval of the ${esc(d.asRefLine)}.`, true) +
    // 4 spaces of hand-fill room right after the label — the T.S.No itself is
    // a manual field the office often hasn't typed in yet when this note is
    // first generated.
    refItem('2)', `T.S.No.&nbsp;&nbsp;&nbsp;&nbsp;${esc(d.tsNo)}/SE/${esc(d.zoneAbbr)}/CMC/${esc(d.financialYear)}, Dated: ${esc(d.tsDate)}`) +
    refItem('3)', `NIT.No.${esc(d.nitNo)}, Dated: ${esc(d.nitDate)} (Item No.${esc(d.itemNo)})`) +
    refItem('4)', `Technical bid opened on ${esc(d.techBidOpenDate)}.`)
  )
}

export function buildBidEvaluationHtml(d: BidEvaluationData): string {
  const rows = d.bidders
    .map(
      (b) =>
        `<tr>${TD(esc(b.sno), 'center')}${TD(esc(b.name))}${TD(
          'The bidder satisfied all tender conditions, if agreed, the bidder will be considered as Responsive'
        )}</tr>`
    )
    .join('')
  return (
    '<p style="margin:0 0 10px;font-weight:bold">Note submitted:-</p>' +
    P(
      `Sub:&emsp;CMC –ENGG– SE – ${esc(d.zoneAbbr)} - ${esc(d.workName)} – Estimate Cost- Rs.${esc(
        d.estimateLakhs
      )} Lakhs - Submission of Technical bid evaluation – Regarding.`
    ) +
    refBlockHtml(d) +
    '<p style="margin:6px 0;text-align:center">***</p>' +
    P(
      `It is to submit that, vide reference 1st cited the ${esc(
        d.asAuthority
      )} has accorded administrative sanction for Rs.${esc(d.estimateLakhs)} Lakhs. Accordingly, Technical Sanction was accorded for the work for Rs.${esc(
        d.estimateLakhs
      )} Lakhs, vide 2nd cited with SSR 2026-27.`
    ) +
    P(
      `In this connection tenders were called vide 3rd cited. The technical bid was opened on e-procurement platform and (${String(d.bidders.length).padStart(
        2,
        '0'
      )}) bidders have participated${
        d.nonRespCount
          ? `, of which (${String(d.nonRespCount).padStart(2, '0')}) bidder(s) were found Non-Responsive and (${String(
              d.bidders.length - d.nonRespCount
            ).padStart(2, '0')}) Responsive`
          : ''
      }`
    ) +
    P(
      `Thus, the technical bid evaluation of the bidder based on the documents uploaded and the hard copies submitted by the bidder is furnished as follows:`
    ) +
    // html-to-docx's own table-cell-width converter only understands
    // pt/px/cm/in (not %) — a % width silently returns undefined and crashes
    // the whole conversion ("Invalid XML name: @w") — so these are in points,
    // sized off the page's 432pt (6in) usable table width.
    `<table style="${TABLE_STYLE}"><tr>${TH('Sl No', '35pt')}${TH('Name of the Bidder', '203pt')}${TH('Remarks', '194pt')}</tr>${rows}</table>` +
    P(
      `The bid evaluation statement of the technical bid is herewith enclosed and submitted for opening of price of responsive bidders.`
    ) +
    SIGN_OFF(d.zoneAbbr, 3)
  )
}

export interface AgencyApprovalData extends SeEvaluationRefs {
  zone: string
  bidEvalApprovedDate: string
  bidders: NoteBidder[]
}

/** The (signed) percentage with its Less/Excess word, and the plain magnitude for display — e.g. "(-)32.89" -> {word:"Less", mag:"32.89"}. */
function pctParts(pct: string): { word: 'Less' | 'Excess'; mag: string } {
  const less = pct.startsWith('(-)')
  return { word: less ? 'Less' : 'Excess', mag: less ? pct.slice(3) : pct }
}

function money(raw: string): string {
  const [whole, frac] = raw.split('.')
  return `${indianDigitGroups(Number(whole))}${frac != null ? '.' + frac : ''}`
}

export function buildAgencyApprovalHtml(d: AgencyApprovalData): string {
  const rows = d.bidders
    .map((b) => {
      const { word, mag } = pctParts(b.pct)
      return `<tr>${TD(esc(b.sno), 'center')}${TD(esc(b.name))}${TD(money(b.ecv), 'right')}${TD(word, 'center')}${TD(
        `${mag}%`,
        'right'
      )}${TD(money(b.tcv), 'right')}${TD(esc(b.rank), 'center')}</tr>`
    })
    .join('')
  const l1 = d.bidders.find((b) => b.rank === 'L-1')
  const l1Pct = l1 ? pctParts(l1.pct) : null
  return (
    '<p style="margin:0 0 10px;font-weight:bold">Note Submitted:</p>' +
    P(
      `Sub:&emsp;CMC- Engg. – SE-${esc(d.zoneAbbr)}- ${esc(d.workName)} – Estimate Cost- Rs.${esc(
        d.estimateLakhs
      )} Lakhs -Evaluation of Financial Bids–Agency approval and issue of  LOA – Regarding.`
    ) +
    refBlockHtml(d) +
    refItem('5)', `Technical Bid evaluation Approved by SE, ${esc(d.zoneAbbr)}, CMC on ${esc(d.bidEvalApprovedDate)}.`) +
    '<p style="margin:6px 0;text-align:center">*********</p>' +
    P(
      `It is to submit that the work “${esc(
        d.workName
      )}” Technical bid evaluation was approved and permitted to open the price bids of eligible tenderer.`
    ) +
    P(`Thus the evaluation of financial bid as follows.`) +
    '<p style="margin:6px 0;font-weight:bold;text-align:center">EVALUATION OF FINANCIAL BID</p>' +
    P(
      `The above work was Technically Sanctioned for Rs.${esc(d.estimateLakhs)} Lakhs by Superintending Engineer, ${esc(
        d.zone
      )} Zone, CMC vide 2nd cited. Tenders were invited through ‘e’ procurement vide 3rd cited and the technical bid opened.`
    ) +
    P(
      `Technical bid was scrutinized in the O/o Superintending Engineer and after careful examination it is found that (${String(
        d.bidders.length
      ).padStart(2, '0')}) bidder were eligible.`
    ) +
    P(`Further, the price bid is opened and the details are furnished below:`) +
    // See the Bid Evaluation table above: pt, not %, for the same
    // html-to-docx compatibility reason — again sized off a 432pt total.
    `<table style="${TABLE_STYLE}"><tr>${TH('Sl.No.', '26pt')}${TH('Contractor/Agency Name', '130pt')}${TH('ECV(INR)', '60pt')}${TH(
      'Excess /Less',
      '43pt'
    )}${TH('Percentage (%)', '48pt')}${TH('Amount (INR)', '65pt')}${TH('Position', '60pt')}</tr>${rows}</table>` +
    (l1 && l1Pct
      ? P(
          `Among the eligible bidders ${esc(l1.name)} is the lowest bidder and quoted @ (-) ${l1Pct.mag} % Less than the ECV with a contract value of Rs. ${money(
            l1.tcv
          )}.`
        )
      : '') +
    P(`Submitted for kind perusal and approval.`) +
    SIGN_OFF(d.zoneAbbr, 2)
  )
}

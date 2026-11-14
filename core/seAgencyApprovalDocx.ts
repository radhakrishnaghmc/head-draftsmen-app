// SE-office Agency Approval note, built straight as OOXML on a strict A4 page (11906 x 16838 twips) so it opens the
// same everywhere. Layout follows the office's own samples (Agency Approval - S VENKANNA / NARENDRA NAIK RAMAVATHU /
// Pusala Sateesh Kumar): the more agencies bid, the smaller the table text, cell padding and line spacing, so the note
// always stays on one A4 sheet instead of spilling its sign-off onto a second page.
import PizZip from 'pizzip'
import type { AgencyApprovalData } from './seEvaluationNotes'
import { indianDigitGroups } from './worksAmounts'

const esc = (s: string): string =>
  String(s ?? '')
    // eslint-disable-next-line no-control-regex
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')

const A4_W = 11906
const A4_H = 16838
const MARGIN_LR = 1440

function money(raw: string): string {
  const [whole, frac] = raw.split('.')
  return `${indianDigitGroups(Number(whole))}${frac != null ? '.' + frac : ''}`
}

function pctParts(pct: string): { word: 'Less' | 'Excess'; mag: string } {
  const less = pct.startsWith('(-)')
  return { word: less ? 'Less' : 'Excess', mag: less ? pct.slice(3) : pct }
}

/** Table text size (half-points), cell padding (twips) and body line spacing for a given number of bidders. */
export function agencyApprovalScale(bidders: number): { size: number; pad: number; line: number } {
  if (bidders <= 3) return { size: 22, pad: 80, line: 348 }
  if (bidders <= 6) return { size: 20, pad: 50, line: 336 }
  if (bidders <= 9) return { size: 18, pad: 30, line: 312 }
  if (bidders <= 13) return { size: 16, pad: 20, line: 288 }
  return { size: 14, pad: 10, line: 264 }
}

const run = (text: string, o: { b?: boolean; size?: number } = {}): string =>
  `<w:r>${o.b || o.size ? `<w:rPr>${o.b ? '<w:b/><w:bCs/>' : ''}${o.size ? `<w:sz w:val="${o.size}"/><w:szCs w:val="${o.size}"/>` : ''}</w:rPr>` : ''}` +
  `<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`

const para = (runs: string, o: { jc?: string; line?: number; after?: number } = {}): string =>
  `<w:p><w:pPr><w:spacing w:after="${o.after ?? 120}" w:line="${o.line ?? 240}" w:lineRule="auto"/>${o.jc ? `<w:jc w:val="${o.jc}"/>` : ''}</w:pPr>${runs}</w:p>`

const NO_BORDERS =
  '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>'
const BOX = (sz: number) =>
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((s) => `<w:${s} w:val="single" w:sz="${sz}" w:space="0" w:color="000000"/>`)
    .join('')

function cell(width: number, content: string): string {
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>${content}</w:tc>`
}

function table(widths: number[], rows: string[], opts: { borders: boolean; pad: number; hPad?: number }): string {
  const hp = opts.hPad ?? 100
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a, b) => a + b, 0)}" w:type="dxa"/><w:jc w:val="center"/>` +
    `${opts.borders ? `<w:tblBorders>${BOX(8)}</w:tblBorders>` : NO_BORDERS}<w:tblLayout w:type="fixed"/>` +
    `<w:tblCellMar><w:top w:w="${opts.pad}" w:type="dxa"/><w:left w:w="${hp}" w:type="dxa"/><w:bottom w:w="${opts.pad}" w:type="dxa"/><w:right w:w="${hp}" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows.join('')}</w:tbl>`
  )
}

const row = (cells: string[]): string => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cells.join('')}</w:tr>`

export function buildAgencyApprovalDocxXml(d: AgencyApprovalData): string {
  const corp = d.corp || 'CMC'
  const n = d.bidders.length
  const { size, pad, line } = agencyApprovalScale(n)
  const J = { jc: 'both', line }
  const IND = '   '

  const refRows: [string, string, string][] = [
    ['Ref:', '1)', `Administrative sanction approval of the ${d.asRefLine}.`],
    ['', '2)', `T.S.No.    ${d.tsNo}/SE/${d.zoneAbbr}/${corp}/${d.financialYear}, Dated: ${d.tsDate}`],
    ['', '3)', `NIT.No.${d.nitNo}, Dated: ${d.nitDate} (Item No.${d.itemNo})`],
    ['', '4)', `Technical bid opened on ${d.techBidOpenDate}.`],
    ['', '5)', `Technical Bid evaluation Approved by SE, ${d.zoneAbbr}, ${corp} on ${d.bidEvalApprovedDate}.`]
  ]
  const refW = [724, 520, 7396]
  const refTable = table(
    refW,
    refRows.map(([label, num, text]) =>
      row([
        cell(refW[0], para(label ? run(label, { b: true }) : '', { after: 40 })),
        cell(refW[1], para(run(num), { after: 40 })),
        cell(refW[2], para(run(text), { jc: 'both', after: 40 }))
      ])
    ),
    { borders: false, pad: 0, hPad: 0 }
  )

  const colW = [800, 1900, 1360, 900, 1260, 1360, 1060]
  const sz = { size }
  const head = [
    'Sl.No.',
    'Contractor/Agency Name',
    'ECV(INR)',
    'Excess /Less',
    'Percentage (%)',
    'Amount (INR)',
    'Position'
  ]
  const bidRows = [
    row(head.map((h, i) => cell(colW[i], para(run(h, { b: true, ...sz }), { jc: 'center', after: 0 })))),
    ...d.bidders.map((b) => {
      const { word, mag } = pctParts(b.pct)
      const vals: [string, string][] = [
        [`${b.sno}`, 'center'],
        [b.name, 'left'],
        [money(b.ecv), 'right'],
        [word, 'center'],
        [`${mag}%`, 'right'],
        [money(b.tcv), 'right'],
        [b.rank, 'center']
      ]
      return row(vals.map(([v, jc], i) => cell(colW[i], para(run(v, sz), { jc, after: 0 }))))
    })
  ]
  const bidTable = table(colW, bidRows, { borders: true, pad })

  const l1 = d.bidders.find((b) => b.rank === 'L-1')
  const l1Pct = l1 ? pctParts(l1.pct) : null
  const count = String(n).padStart(2, '0')
  const signW = [1725, 2583, 4332]
  const sign = table(
    signW,
    [
      row([
        cell(signW[0], para(run('AE', { b: true }), { after: 0 })),
        cell(signW[1], para(run('Dy.EE', { b: true }), { jc: 'center', after: 0 })),
        cell(signW[2], para(run(`SE (M), ${d.zoneAbbr}`, { b: true }), { jc: 'right', after: 0 }))
      ])
    ],
    { borders: false, pad: 80, hPad: 160 }
  )

  const body = [
    para(run('Note Submitted:', { b: true })),
    para(
      run(
        `${IND}Sub: ${corp}- Engg. – SE-${d.zoneAbbr}- ${d.workName} – Estimate Cost- Rs.${d.estimateLakhs} Lakhs -Evaluation of Financial Bids–Agency approval and issue of  LOA – Regarding.`
      ),
      J
    ),
    refTable,
    para('', { after: 0 }),
    para(run('*********'), { jc: 'center' }),
    para(
      run(
        `${IND}It is to submit that the work “${d.workName}” and Technically Sanctioned for Rs.${d.estimateLakhs} Lakhs by the Superintending Engineer, ${d.zone} Zone, ${corp} vide 2nd cited`
      ),
      J
    ),
    para(
      run(
        `${IND}Tenders were invited through ‘e’ procurement vide 3rd cited and the technical bid opened. Technical bid was scrutinized in the O/o Superintending Engineer and found that (${count}) ${
          n === 1 ? 'bidder was' : 'bidders were'
        } eligible.`
      ),
      J
    ),
    para(run(`${IND}Further, the price bid is opened and the details are furnished below:`), J),
    bidTable,
    para('', { after: 0 }),
    l1 && l1Pct
      ? para(
          run(
            `${IND}Among the eligible bidders ${l1.name} is the lowest bidder and quoted @ (-) ${l1Pct.mag} % ${l1Pct.word} than the ECV with a contract value of Rs. ${money(l1.tcv)}.`
          ),
          J
        )
      : '',
    para(run(`${IND}Submitted for kind perusal and approval.`), J),
    para('', J),
    para('', J),
    sign
  ].join('')

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}` +
    `<w:p/><w:sectPr><w:pgSz w:w="${A4_W}" w:h="${A4_H}" w:code="9"/>` +
    `<w:pgMar w:top="1134" w:right="${MARGIN_LR}" w:bottom="1134" w:left="${MARGIN_LR}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`
  )
}

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
const PACKAGE_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
const DOC_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>'

/** The Agency Approval note as a base64 .docx (A4). */
export function buildAgencyApprovalDocxB64(d: AgencyApprovalData): string {
  const zip = new PizZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('word/document.xml', buildAgencyApprovalDocxXml(d))
  zip.file('word/_rels/document.xml.rels', DOC_RELS)
  zip.file('word/styles.xml', STYLES)
  return zip.generate({ type: 'base64', compression: 'DEFLATE' })
}

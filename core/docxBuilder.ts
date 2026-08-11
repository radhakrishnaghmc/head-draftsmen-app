import PizZip from 'pizzip'

/**
 * Build a Word (.docx) file DIRECTLY as clean OOXML, rather than through
 * html-to-docx (whose output Microsoft Word refuses to open — misplaced body
 * section properties, package directory entries, and other quirks that only
 * Word rejects). Every part here is minimal, correctly-ordered OOXML modelled on
 * what Word/python-docx emit, so the result opens cleanly in Word.
 *
 * Used by the Photos/PDF → Word tool to write the reconstructed layout (real
 * paragraphs + tables) as an editable document.
 */

export interface DocRun {
  text: string
  bold?: boolean
  /** Font size in half-points (e.g. 22 = 11pt). Omitted → the document default. */
  size?: number
}
export interface DocParagraph {
  kind: 'paragraph'
  runs: DocRun[]
  align?: 'left' | 'center'
  /** A hard page break before this paragraph's content (used between pages). */
  pageBreak?: boolean
}
export interface DocCell {
  /** Each inner array is one line (paragraph) of runs inside the cell. */
  lines: DocRun[][]
}
export interface DocTable {
  kind: 'table'
  /** Column widths in twips (1/20 pt); their sum is the table width. */
  colWidths: number[]
  rows: DocCell[][]
}
export type DocBlock = DocParagraph | DocTable

const esc = (s: string): string =>
  String(s ?? '')
    // Strip characters XML 1.0 forbids (control chars other than tab/LF/CR).
    // PDF/OCR text can carry these (e.g. 0x02); left in, they make document.xml
    // invalid XML and Microsoft Word refuses to open the file, even though
    // LibreOffice tolerates it. This must run before entity-escaping.
    // eslint-disable-next-line no-control-regex
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function runXml(run: DocRun): string {
  const props: string[] = []
  if (run.bold) props.push('<w:b/><w:bCs/>')
  if (run.size) props.push(`<w:sz w:val="${Math.round(run.size)}"/><w:szCs w:val="${Math.round(run.size)}"/>`)
  const rPr = props.length ? `<w:rPr>${props.join('')}</w:rPr>` : ''
  return `<w:r>${rPr}<w:t xml:space="preserve">${esc(run.text)}</w:t></w:r>`
}

function paragraphXml(p: DocParagraph): string {
  const props: string[] = []
  if (p.align === 'center') props.push('<w:jc w:val="center"/>')
  const pPr = props.length ? `<w:pPr>${props.join('')}</w:pPr>` : ''
  const brk = p.pageBreak ? '<w:r><w:br w:type="page"/></w:r>' : ''
  const runs = p.runs.map(runXml).join('')
  return `<w:p>${pPr}${brk}${runs}</w:p>`
}

/** Cell content: one <w:p> per line; an empty cell still needs a single <w:p/>. */
function cellXml(cell: DocCell, width: number): string {
  const tcPr = `<w:tcPr><w:tcW w:type="dxa" w:w="${Math.round(width)}"/></w:tcPr>`
  const lines = cell.lines.filter((l) => l.length > 0)
  const body = lines.length
    ? lines.map((line) => `<w:p>${line.map(runXml).join('')}</w:p>`).join('')
    : '<w:p/>'
  return `<w:tc>${tcPr}${body}</w:tc>`
}

const BORDER = (v: string) => `<w:${v} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`

function tableXml(t: DocTable): string {
  const total = t.colWidths.reduce((s, w) => s + w, 0)
  const tblPr =
    `<w:tblPr>` +
    `<w:tblW w:type="dxa" w:w="${Math.round(total)}"/>` +
    `<w:tblBorders>${BORDER('top')}${BORDER('left')}${BORDER('bottom')}${BORDER('right')}${BORDER('insideH')}${BORDER('insideV')}</w:tblBorders>` +
    `</w:tblPr>`
  const grid = `<w:tblGrid>${t.colWidths.map((w) => `<w:gridCol w:w="${Math.round(w)}"/>`).join('')}</w:tblGrid>`
  const rows = t.rows
    .map((row) => `<w:tr>${row.map((cell, c) => cellXml(cell, t.colWidths[c] ?? total / row.length)).join('')}</w:tr>`)
    .join('')
  return `<w:tbl>${tblPr}${grid}${rows}</w:tbl>`
}

function blockXml(block: DocBlock): string {
  return block.kind === 'table' ? tableXml(block) : paragraphXml(block)
}

function documentXml(blocks: DocBlock[]): string {
  const body = blocks.map(blockXml).join('')
  // The body-level sectPr MUST be the last child (letter page, 1" margins).
  const sectPr =
    `<w:sectPr>` +
    `<w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr>`
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W_NS}"><w:body>${body}${sectPr}</w:body></w:document>`
  )
}

const CONTENT_TYPES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
  `</Types>`

const PACKAGE_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`

const DOCUMENT_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
  `</Relationships>`

const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:styles xmlns:w="${W_NS}">` +
  `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  `</w:styles>`

/**
 * Assemble the doc-model blocks into a complete, Word-valid .docx buffer. PizZip
 * (which builds via explicit file entries) produces a package with no directory
 * entries — one of the things that made the html-to-docx output unopenable.
 */
/** Flatten doc-model blocks into a 2-D grid for Excel: a table's cells become rows; a paragraph becomes a single-cell row. */
export function blocksToRows(blocks: DocBlock[]): string[][] {
  const rows: string[][] = []
  for (const b of blocks) {
    if (b.kind === 'table') {
      for (const row of b.rows) {
        rows.push(row.map((cell) => cell.lines.map((line) => line.map((r) => r.text).join('')).join(' ').trim()))
      }
    } else {
      const text = b.runs.map((r) => r.text).join('').trim()
      if (text) rows.push([text])
    }
  }
  return rows.length ? rows : [['']]
}

export function buildDocx(blocks: DocBlock[]): Buffer {
  const zip = new PizZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('word/document.xml', documentXml(blocks))
  zip.file('word/_rels/document.xml.rels', DOCUMENT_RELS)
  zip.file('word/styles.xml', STYLES)
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as Buffer
}

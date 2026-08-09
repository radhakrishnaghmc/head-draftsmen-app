import PizZip from 'pizzip'

// A body-level <w:sectPr> in either form, anchored at the very end (Word /
// LibreOffice) or the very start (html-to-docx) of the body content.
const TRAILING_SECTPR = /(?:<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>)\s*$/
const LEADING_SECTPR = /^\s*(?:<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>)/

// Top-level body children, matched atomically: a table, a content control, a
// self-closing empty paragraph, or a paragraph. (Nested tables/sdt — rare in the
// app's documents — aren't handled by the non-greedy match, which is acceptable.)
const BLOCK_RE = /<w:tbl[\s\S]*?<\/w:tbl>|<w:sdt[\s\S]*?<\/w:sdt>|<w:p\b[^>]*\/>|<w:p\b[\s\S]*?<\/w:p>/g

// A manual page break (<w:br w:type="page"/>) anywhere inside a block.
const PAGE_BREAK_IN = /<w:br\b[^>]*w:type="page"/
// A "page break before" paragraph property that is on (not w:val="false/0/off").
const PAGE_BREAK_BEFORE = /<w:pageBreakBefore\b/
const PAGE_BREAK_BEFORE_OFF = /<w:pageBreakBefore\b[^>]*w:val="(?:0|false|off)"/

/**
 * Split a .docx into one .docx per page — cutting at manual page breaks
 * (`<w:br w:type="page"/>`) and "page break before" paragraphs. Each piece is a
 * full clone of the original document (its styles, numbering, headers/footers,
 * images and relationships are all kept) with only the body content narrowed to
 * that page's blocks, so every output opens as a proper, well-formatted Word
 * file. A paragraph is the smallest unit — a break splits between paragraphs, not
 * within one. Returns one buffer when the document has no page breaks (i.e. the
 * original, unchanged).
 */
export function splitDocxByPageBreaks(buffer: Buffer): Buffer[] {
  const zip = new PizZip(buffer)
  const file = zip.file('word/document.xml')
  if (!file) throw new Error('Not a valid .docx (no word/document.xml).')
  const xml = file.asText()

  const openTag = xml.search(/<w:body[\s>]/)
  if (openTag === -1) return [buffer]
  const openEnd = xml.indexOf('>', openTag) + 1
  const close = xml.lastIndexOf('</w:body>')
  if (close === -1) return [buffer]

  const head = xml.slice(0, openEnd) // …<w:body>
  const tail = xml.slice(close) // </w:body></w:document>
  let inner = xml.slice(openEnd, close)

  // Lift the body-level sectPr (page setup) out so it isn't treated as content;
  // it's reattached at the end of every piece.
  let bodySect = ''
  let m = inner.match(TRAILING_SECTPR)
  if (m) {
    bodySect = m[0].trim()
    inner = inner.slice(0, m.index)
  } else if ((m = inner.match(LEADING_SECTPR))) {
    bodySect = m[0].trim()
    inner = inner.slice(m[0].length)
  }

  const blocks = inner.match(BLOCK_RE) ?? []
  if (blocks.length === 0) return [buffer]

  // Group blocks into pages: a "page break before" starts a new page ahead of its
  // block; a manual page break ends the page after its block.
  const pages: string[][] = [[]]
  for (const b of blocks) {
    const breakBefore = PAGE_BREAK_BEFORE.test(b) && !PAGE_BREAK_BEFORE_OFF.test(b)
    if (breakBefore && pages[pages.length - 1].length > 0) pages.push([])
    pages[pages.length - 1].push(b)
    if (PAGE_BREAK_IN.test(b)) pages.push([])
  }
  const groups = pages.filter((g) => g.length > 0)
  if (groups.length <= 1) return [buffer]

  return groups.map((g) => {
    // Clone the whole original archive so each piece keeps every style, header,
    // image and relationship; only the body is narrowed to this page's blocks.
    const z = new PizZip(buffer)
    z.file('word/document.xml', head + g.join('') + bodySect + tail)
    return z.generate({ type: 'nodebuffer' })
  })
}

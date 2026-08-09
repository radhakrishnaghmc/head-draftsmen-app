import PizZip from 'pizzip'

// A page break wrapped in its own paragraph — inserted between two documents so
// each merged file starts on a fresh page.
const PAGE_BREAK_P = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

// Section properties, in every form they appear: a paired <w:sectPr>…</w:sectPr>
// or a self-closing <w:sectPr/>. Different producers put the body-level sectPr in
// different places — Word/LibreOffice at the very end of the body, html-to-docx
// at the very start — and paragraphs may carry their own section-break sectPr.
const SECTPR_RE = /<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/g

/**
 * The inner content of a document's <w:body> with all <w:sectPr> removed. We
 * drop every sectPr (body-level page setup, and any paragraph-level section
 * break) so an appended document contributes only its paragraphs and tables —
 * the base document's own section properties govern the merged file. Stripping a
 * paragraph's sectPr only removes a section break; its text is untouched.
 */
function bodyInner(documentXml: string): string {
  const openTag = documentXml.search(/<w:body[\s>]/)
  if (openTag === -1) return ''
  const openEnd = documentXml.indexOf('>', openTag) + 1
  const close = documentXml.lastIndexOf('</w:body>')
  if (close === -1) return ''
  return documentXml.slice(openEnd, close).replace(SECTPR_RE, '')
}

/**
 * Where to splice appended content into the base body: just before a *trailing*
 * body-level sectPr when there is one (Word/LibreOffice — the sectPr must remain
 * the last child of the body), otherwise right before </w:body> (html-to-docx,
 * whose sectPr sits at the start of the body).
 */
function baseInsertPos(baseXml: string): number {
  const bodyClose = baseXml.lastIndexOf('</w:body>')
  if (bodyClose === -1) return baseXml.length
  const before = baseXml.slice(0, bodyClose)
  const lastSectClose = before.lastIndexOf('</w:sectPr>')
  if (lastSectClose !== -1 && before.slice(lastSectClose + '</w:sectPr>'.length).trim() === '') {
    // The sectPr is the body's last child — insert before it, not after.
    const open = before.lastIndexOf('<w:sectPr')
    if (open !== -1) return open
  }
  return bodyClose
}

/**
 * Merge several .docx files, in order, into one .docx. The first document is the
 * base — its styles, numbering, headers/footers and page setup are what the
 * merged file keeps — and every later document's body is appended after a page
 * break. This is a text/table-oriented merge (the app's own generated documents):
 * fonts, paragraph and table content carry over, but numbering restarts and
 * images/relationships defined only in a later file may not survive, since those
 * live in parts that aren't remapped. Returns the merged .docx as a Buffer.
 */
export function mergeDocxBuffers(buffers: Buffer[]): Buffer {
  if (buffers.length === 0) throw new Error('No documents to merge.')
  if (buffers.length === 1) return buffers[0]

  const baseZip = new PizZip(buffers[0])
  const baseFile = baseZip.file('word/document.xml')
  if (!baseFile) throw new Error('The first file is not a valid .docx (no word/document.xml).')
  const baseXml = baseFile.asText()

  let insertion = ''
  for (const buf of buffers.slice(1)) {
    const file = new PizZip(buf).file('word/document.xml')
    if (!file) continue
    insertion += PAGE_BREAK_P + bodyInner(file.asText())
  }

  const at = baseInsertPos(baseXml)
  const mergedXml = baseXml.slice(0, at) + insertion + baseXml.slice(at)
  baseZip.file('word/document.xml', mergedXml)
  return baseZip.generate({ type: 'nodebuffer' })
}

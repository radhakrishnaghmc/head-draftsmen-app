import PizZip from 'pizzip'

/** Smallest run size ever written, in half-points (8pt) — below this a printed page is unreadable. */
const MIN_HALF_POINTS = 16

const scaleSizes = (xml: string, factor: number): string =>
  xml
    .replace(
      /<w:(sz|szCs) w:val="(\d+)"/g,
      (_m, tag: string, v: string) => `<w:${tag} w:val="${Math.max(MIN_HALF_POINTS, Math.round(Number(v) * factor))}"`
    )
    .replace(
      /(<w:spacing\b[^>]*?\sw:(?:before|after)=")(\d+)(")/g,
      (_m, a: string, v: string, c: string) => `${a}${Math.round(Number(v) * factor)}${c}`
    )

/**
 * Word sizes text with no size anywhere (not in the run, the style, or docDefaults) at 10pt; Chromium
 * would use its own 12pt, printing such a document ~20% larger than Word shows it. Writes the 10pt in.
 */
function withDefaultSize(stylesXml: string): string {
  const normal = /<w:style\b[^>]*w:default="1"[^>]*>[\s\S]*?<\/w:style>/.exec(stylesXml)?.[0] ?? ''
  const defaults = /<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/.exec(stylesXml)?.[0] ?? ''
  if (/<w:sz\b/.test(normal) || /<w:sz\b/.test(defaults)) return stylesXml
  return stylesXml.replace(/<w:rPrDefault><w:rPr>([\s\S]*?)<\/w:rPr>/, (_m, inner: string) => {
    const size = '<w:sz w:val="20"/><w:szCs w:val="20"/>'
    const at = inner.indexOf('<w:lang')
    return `<w:rPrDefault><w:rPr>${at < 0 ? inner + size : inner.slice(0, at) + size + inner.slice(at)}</w:rPr>`
  })
}

/**
 * Shrinks a base64 .docx's body text and paragraph spacing by `factor` (0–1; 1 only normalises the default size) so a page that just overflows A4
 * fits on one. Blank paragraphs above the first line of text are left alone — on the bond templates they are the fixed gap for the
 * pre-printed letterhead — as are paragraphs holding a drawing/text box (the logo banner). The style sheet is
 * scaled too, since body text with no explicit size inherits its size from it.
 */
export function scaleDocxText(docxB64: string, factor: number): string {
  const zip = new PizZip(docxB64, { base64: true })
  const doc = zip.file('word/document.xml')
  if (!doc) return docxB64
  let seenText = false
  const xml =
    factor >= 1
      ? doc.asText()
      : doc.asText().replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (p) => {
          const hasText = /<w:t(?:\s[^>]*)?>[^<]+<\/w:t>/.test(p)
          if (/<w:drawing|<w:pict|<mc:AlternateContent/.test(p)) return p
          // Blank lines above the title are the letterhead gap; blank spacer lines below it can shrink with the text.
          if (!hasText && !seenText) return p
          if (hasText) seenText = true
          return scaleSizes(p, factor)
        })
  zip.file('word/document.xml', xml)
  const styles = zip.file('word/styles.xml')
  if (styles) {
    const withSize = withDefaultSize(styles.asText())
    zip.file('word/styles.xml', factor < 1 ? scaleSizes(withSize, factor) : withSize)
  }
  return zip.generate({ type: 'base64', compression: 'DEFLATE' })
}

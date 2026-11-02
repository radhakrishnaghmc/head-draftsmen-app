import PizZip from 'pizzip'

export interface PageMargins {
  top: number
  bottom: number
  left: number
  right: number
}

const TWIPS_PER_MM = 1440 / 25.4
const A4_WIDTH_TWIPS = 11906

/**
 * Rewrites every section's page margins (<w:pgMar>) in a base64 .docx (header/footer/gutter distances are kept),
 * and/or — with `a4` — its page size to A4 portrait. Pass `m` null to leave margins alone.
 */
export function applyPageMargins(docxB64: string, m: PageMargins | null, a4 = false): string {
  const zip = new PizZip(docxB64, { base64: true })
  const file = zip.file('word/document.xml')
  if (!file) return docxB64
  const tw = (mm: number): string => String(Math.round(mm * TWIPS_PER_MM))
  let xml = file.asText()
  if (a4) xml = xml.replace(/<w:pgSz\b[^>]*\/>/g, '<w:pgSz w:w="11906" w:h="16838"/>')
  // The A4-forced templates were laid out for US Letter, so their tables (and the logo banner inside
  // them) can be wider than A4's text area and run past the right margin. With no hand-set margins,
  // narrow the left/right margins evenly just enough for the widest table to sit inside, centred.
  if (a4 && !m) {
    const widest = Math.max(0, ...[...xml.matchAll(/<w:tblW w:w="(\d+)" w:type="dxa"\/>/g)].map((t) => Number(t[1])))
    if (widest > 0 && widest <= A4_WIDTH_TWIPS) {
      xml = xml.replace(/<w:pgMar\b[^>]*\/>/g, (tag) => {
        const side = Math.floor((A4_WIDTH_TWIPS - widest) / 2)
        const left = Number(/w:left="(\d+)"/.exec(tag)?.[1] ?? 0)
        const right = Number(/w:right="(\d+)"/.exec(tag)?.[1] ?? 0)
        if (A4_WIDTH_TWIPS - left - right >= widest) return tag
        return tag.replace(/w:left="[^"]*"/, `w:left="${side}"`).replace(/w:right="[^"]*"/, `w:right="${side}"`)
      })
    }
  }
  if (m) xml = xml.replace(/<w:pgMar\b[^>]*\/>/g, (tag) =>
    tag
      .replace(/w:top="[^"]*"/, `w:top="${tw(m.top)}"`)
      .replace(/w:bottom="[^"]*"/, `w:bottom="${tw(m.bottom)}"`)
      .replace(/w:left="[^"]*"/, `w:left="${tw(m.left)}"`)
      .replace(/w:right="[^"]*"/, `w:right="${tw(m.right)}"`)
  )
  zip.file('word/document.xml', xml)
  return zip.generate({ type: 'base64', compression: 'DEFLATE' })
}

/**
 * SE (Zone-level) File Backer header: trims the blank spacer lines around the corporation banner (about three lines in
 * all) and sets the office line below it at 14pt so "Office of Superintending Engineer, <Zone> Zone, <Corp>" stays on
 * one line. The EE templates keep their own spacing — this is applied to the SE File Backer only.
 */
export function tightenSeFileBackerHeader(docxB64: string): string {
  const zip = new PizZip(docxB64, { base64: true })
  const file = zip.file('word/document.xml')
  if (!file) return docxB64
  let xml = file.asText()
  // The top spacer paragraph above the banner (24pt) and the empty 20pt paragraph between banner and office line.
  xml = xml.replace('<w:sz w:val="48"/><w:szCs w:val="48"/>', '<w:sz w:val="8"/><w:szCs w:val="8"/>')
  xml = xml.replace(
    '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:sz w:val="40"/><w:szCs w:val="36"/></w:rPr></w:pPr></w:p>',
    '<w:pPr><w:jc w:val="center"/><w:rPr><w:b/><w:sz w:val="8"/><w:szCs w:val="8"/></w:rPr></w:pPr></w:p>'
  )
  xml = xml.replace(
    '<w:rPr><w:b/><w:sz w:val="40"/><w:szCs w:val="36"/></w:rPr><w:t>{{Circle Header}}',
    '<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>{{Circle Header}}'
  )
  zip.file('word/document.xml', xml)
  return zip.generate({ type: 'base64', compression: 'DEFLATE' })
}

/** Deletes the first paragraph whose text contains `needle` (e.g. an optional line of a template). No-op if absent. */
export function removeDocParagraph(docxB64: string, needle: string): string {
  const zip = new PizZip(docxB64, { base64: true })
  const file = zip.file('word/document.xml')
  if (!file) return docxB64
  const xml = file.asText()
  const para = /<w:p(?: [^>]*)?>(?:(?!<w:p[ >]).)*?<\/w:p>/gs
  let hit: RegExpExecArray | null
  while ((hit = para.exec(xml))) {
    const text = [...hit[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('')
    if (text.includes(needle)) {
      zip.file('word/document.xml', xml.slice(0, hit.index) + xml.slice(hit.index + hit[0].length))
      return zip.generate({ type: 'base64', compression: 'DEFLATE' })
    }
  }
  return docxB64
}

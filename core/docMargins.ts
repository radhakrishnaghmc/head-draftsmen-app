import PizZip from 'pizzip'

export interface PageMargins {
  top: number
  bottom: number
  left: number
  right: number
}

const TWIPS_PER_MM = 1440 / 25.4

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

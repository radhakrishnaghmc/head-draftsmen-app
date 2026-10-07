import HTMLtoDOCX from 'html-to-docx'
import { sanitizeHtmlDocxForWord2007 } from './word2007Compat'

/** html-to-docx's own portrait defaults (twips), so a caller overriding one
 * margin doesn't blank out the rest. */
const DEFAULT_MARGINS = { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 }

/** Convert filled document HTML into a real .docx file, formatting intact.
 * The output is passed through sanitizeHtmlDocxForWord2007 so it opens in Word
 * 2007 (html-to-docx emits invalid .rels content-types and out-of-order
 * property children that only Word 2010+ tolerates). */
export async function convertHtmlToDocx(
  html: string,
  margins?: Partial<typeof DEFAULT_MARGINS>
): Promise<Buffer> {
  const body = /<html/i.test(html) ? html : `<!DOCTYPE html><html><body>${html}</body></html>`
  const result = await HTMLtoDOCX(body, undefined, {
    table: { row: { cantSplit: true } },
    footer: false,
    header: false,
    margins: margins ? { ...DEFAULT_MARGINS, ...margins } : undefined
  })
  const buffer = Buffer.isBuffer(result) ? result : Buffer.from(result as ArrayBuffer)
  return sanitizeHtmlDocxForWord2007(buffer)
}

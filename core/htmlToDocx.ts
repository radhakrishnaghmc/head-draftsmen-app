import HTMLtoDOCX from 'html-to-docx'
import { sanitizeHtmlDocxForWord2007 } from './word2007Compat'

/** Convert filled document HTML into a real .docx file, formatting intact.
 * The output is passed through sanitizeHtmlDocxForWord2007 so it opens in Word
 * 2007 (html-to-docx emits invalid .rels content-types and out-of-order
 * property children that only Word 2010+ tolerates). */
export async function convertHtmlToDocx(html: string): Promise<Buffer> {
  const body = /<html/i.test(html) ? html : `<!DOCTYPE html><html><body>${html}</body></html>`
  const result = await HTMLtoDOCX(body, undefined, {
    table: { row: { cantSplit: true } },
    footer: false,
    header: false
  })
  const buffer = Buffer.isBuffer(result) ? result : Buffer.from(result as ArrayBuffer)
  return sanitizeHtmlDocxForWord2007(buffer)
}

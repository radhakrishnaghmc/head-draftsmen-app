import HTMLtoDOCX from 'html-to-docx'

/** Convert filled document HTML into a real .docx file, formatting intact. */
export async function convertHtmlToDocx(html: string): Promise<Buffer> {
  const body = /<html/i.test(html) ? html : `<!DOCTYPE html><html><body>${html}</body></html>`
  const result = await HTMLtoDOCX(body, undefined, {
    table: { row: { cantSplit: true } },
    footer: false,
    header: false
  })
  return Buffer.isBuffer(result) ? result : Buffer.from(result as ArrayBuffer)
}

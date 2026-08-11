import libre from 'libreoffice-convert'
import { promisify } from 'util'

// libreoffice-convert only sets the OUTPUT filter; a PDF has to be opened with
// the INPUT filter `writer_pdf_import`, otherwise LibreOffice loads it in Draw
// (which cannot export .docx) and the conversion silently produces nothing. The
// input filter is passed straight through to `soffice` via sofficeAdditionalArgs,
// and fileName gives the temp input a `.pdf` extension so the output is named
// `source.docx`.
const convertWithOptions = promisify(libre.convertWithOptions) as (
  input: Buffer,
  outputExt: string,
  filter: string | undefined,
  options: { fileName?: string; sofficeAdditionalArgs?: string[] }
) => Promise<Buffer>

/**
 * Convert a PDF buffer to a .docx that preserves the PDF's layout — the text,
 * tables, borders and positioning come across as real, editable Word content
 * (via LibreOffice's `writer_pdf_import`), rather than the plain OCR'd word
 * stream that loses all formatting. Best for text-based PDFs; a scanned/image
 * PDF still converts, but its pages come in as images (use the OCR "text" path
 * when you need editable text out of a scan).
 *
 * Retries a few times: back-to-back conversions can transiently fail while a
 * prior `soffice` instance is still releasing its user-profile lock.
 */
export async function convertPdfToDocx(pdfBuffer: Buffer): Promise<Buffer> {
  let lastErr: unknown
  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      return await convertWithOptions(pdfBuffer, 'docx', undefined, {
        fileName: 'source.pdf',
        sofficeAdditionalArgs: ['--infilter=writer_pdf_import']
      })
    } catch (err) {
      lastErr = err
      await new Promise((resolve) => setTimeout(resolve, 700 * (attempt + 1)))
    }
  }
  throw new Error(
    'Layout-preserving PDF → Word needs LibreOffice installed on this computer. ' +
      'Install it from libreoffice.org, then try again — or use the OCR "text" option instead.\n' +
      (lastErr instanceof Error ? lastErr.message : String(lastErr))
  )
}

import libre from 'libreoffice-convert'
import { promisify } from 'util'

const convertAsync = promisify(libre.convert) as (
  input: Buffer,
  outputExt: string,
  filter: string | undefined
) => Promise<Buffer>

/**
 * Convert a .docx buffer to PDF via a local LibreOffice install (the
 * `libreoffice-convert` package shells out to `soffice`). Throws a clear
 * error when LibreOffice isn't installed, instead of a cryptic ENOENT.
 */
export async function convertDocxToPdf(docxBuffer: Buffer): Promise<Buffer> {
  try {
    return await convertAsync(docxBuffer, '.pdf', undefined)
  } catch (err) {
    throw new Error(
      'PDF conversion requires LibreOffice to be installed on this computer. ' +
        'Install it from libreoffice.org, then try again — or use the Word (.docx) or Print option instead.\n' +
        (err instanceof Error ? err.message : String(err))
    )
  }
}

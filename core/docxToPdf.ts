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
 *
 * Retries a few times before giving up: when several documents are converted
 * back-to-back (the "Download all" bundle converts 3 PDFs in a row), a prior
 * conversion's `soffice` instance can still be shutting down and holding the
 * user-profile lock, so the next call fails transiently. A short wait-and-retry
 * lets that instance exit — without it, the later PDFs in a bundle silently
 * dropped out.
 */
export async function convertDocxToPdf(docxBuffer: Buffer): Promise<Buffer> {
  let lastErr: unknown
  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      return await convertAsync(docxBuffer, '.pdf', undefined)
    } catch (err) {
      lastErr = err
      // Wait a beat (growing) for a lingering soffice instance to release its
      // profile lock, then retry.
      await new Promise((resolve) => setTimeout(resolve, 700 * (attempt + 1)))
    }
  }
  throw new Error(
    'PDF conversion requires LibreOffice to be installed on this computer. ' +
      'Install it from libreoffice.org, then try again — or use the Word (.docx) or Print option instead.\n' +
      (lastErr instanceof Error ? lastErr.message : String(lastErr))
  )
}

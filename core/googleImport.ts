import * as http from 'http'
import * as https from 'https'
import type { ExcelTable } from './types'
import { parseExcelBuffer, allSheetGridsFromBuffer } from './excel'
import type { SheetGrid } from './sheet'

const USER_AGENT =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'

/**
 * Turn a Google Sheets / Drive share link into a direct-download URL for the
 * raw spreadsheet bytes. Links that don't match either pattern are returned
 * as-is (e.g. an already-direct download URL).
 *
 * A copied Google Sheets URL almost always carries a `#gid=` fragment for
 * whichever tab happened to be open at the time — respected here so a link
 * to one specific sheet exports just that sheet. Pass `wholeWorkbook: true`
 * to ignore it and export every sheet instead (see
 * importAllSheetsFromGoogleLink, where honoring a stray gid would silently
 * limit a multi-sheet monitoring workbook down to whichever one tab was
 * active when the link was copied).
 */
export function resolveGoogleDownloadUrl(link: string, opts?: { wholeWorkbook?: boolean }): string {
  const url = link.trim()

  const sheetsMatch = url.match(/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/)
  if (sheetsMatch) {
    const id = sheetsMatch[1]
    const gidMatch = opts?.wholeWorkbook ? null : url.match(/[?#&]gid=(\d+)/)
    const gid = gidMatch ? gidMatch[1] : null
    return `https://docs.google.com/spreadsheets/d/${id}/export?format=xlsx${gid ? `&gid=${gid}` : ''}`
  }

  if (url.includes('drive.google.com')) {
    const driveMatch = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/)
    if (driveMatch) return `https://drive.google.com/uc?export=download&id=${driveMatch[1]}`
  }

  return url
}

/** Download a binary resource, following redirects. */
function download(url: string, redirectsLeft = 5): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const client = url.startsWith('http://') ? http : https
    const req = client.get(url, { headers: { 'user-agent': USER_AGENT } }, (res) => {
      const { statusCode, headers } = res
      if (statusCode && statusCode >= 300 && statusCode < 400 && headers.location) {
        res.resume()
        if (redirectsLeft <= 0) {
          reject(new Error('Too many redirects while downloading the link.'))
          return
        }
        const next = new URL(headers.location, url).toString()
        download(next, redirectsLeft - 1).then(resolve, reject)
        return
      }
      if (statusCode && statusCode >= 400) {
        reject(
          new Error(
            `The link returned an error (HTTP ${statusCode}). Make sure it's shared as "Anyone with the link can view".`
          )
        )
        res.resume()
        return
      }
      const chunks: Buffer[] = []
      res.on('data', (c) => chunks.push(c))
      res.on('end', () => resolve(Buffer.concat(chunks)))
    })
    req.on('error', reject)
    req.setTimeout(20000, () => req.destroy(new Error('Timed out downloading the link.')))
  })
}

/** Download a Google Sheets / Drive spreadsheet link and parse it into an ExcelTable. */
export async function importTableFromGoogleLink(link: string): Promise<ExcelTable> {
  const downloadUrl = resolveGoogleDownloadUrl(link)
  const buffer = await download(downloadUrl)

  let table: ExcelTable
  try {
    table = parseExcelBuffer(buffer, 'Imported from link', link)
  } catch {
    throw new Error(
      'Could not read spreadsheet data from that link. Make sure it points to a Google Sheet or Excel file shared as "Anyone with the link can view".'
    )
  }
  if (table.headers.length === 0 || table.rows.length === 0) {
    throw new Error(
      'No data found at that link. Make sure it points to a Google Sheet or Excel file shared as "Anyone with the link can view".'
    )
  }
  return table
}

/**
 * Download a Google Sheets link and parse *every* sheet in the workbook —
 * used for a "monitoring format" workbook that holds one sheet per circle,
 * unlike the single-sheet Works List import above. Always fetches the whole
 * workbook (ignoring any #gid= in the pasted link) since the point is to
 * search across every circle's sheet, not just whichever tab was open when
 * the link was copied.
 */
export async function importAllSheetsFromGoogleLink(link: string): Promise<SheetGrid[]> {
  const downloadUrl = resolveGoogleDownloadUrl(link, { wholeWorkbook: true })
  const buffer = await download(downloadUrl)

  let sheets: SheetGrid[]
  try {
    sheets = allSheetGridsFromBuffer(buffer, 'Monitoring format', link)
  } catch {
    throw new Error(
      'Could not read spreadsheet data from that link. Make sure it points to a Google Sheet or Excel file shared as "Anyone with the link can view".'
    )
  }
  if (sheets.length === 0) {
    throw new Error(
      'No sheets found at that link. Make sure it points to a Google Sheet or Excel file shared as "Anyone with the link can view".'
    )
  }
  return sheets
}

import { download } from './googleImport'

export interface DriveFile {
  id: string
  name: string
}

/** Extract the folder ID from a Google Drive folder share link, or null if the link isn't a Drive folder link. */
export function driveFolderId(link: string): string | null {
  const m =
    link.match(/drive\.google\.com\/drive\/folders\/([a-zA-Z0-9_-]+)/) ||
    link.match(/drive\.google\.com\/drive\/u\/\d+\/folders\/([a-zA-Z0-9_-]+)/) ||
    link.match(/[?&]id=([a-zA-Z0-9_-]+)/)
  return m ? m[1] : null
}

/**
 * Parse Google Drive's public "embeddedfolderview" listing HTML into its
 * files. That view (drive.google.com/embeddedfolderview?id=…#list) is the
 * stable, no-API-key way to enumerate a publicly shared folder: each file is
 * a `<div class="flip-entry" id="entry-<FILE_ID>">` whose
 * `<div class="flip-entry-title">` holds the file name. Deduped by id, in
 * listing order.
 */
export function parseDriveFolderListing(html: string): DriveFile[] {
  const files: DriveFile[] = []
  const seen = new Set<string>()
  const re = /id="entry-([a-zA-Z0-9_-]+)"[\s\S]*?flip-entry-title[^>]*>([^<]+)</g
  let m: RegExpExecArray | null
  while ((m = re.exec(html)) !== null) {
    const id = m[1]
    const name = m[2].trim()
    if (id && !seen.has(id)) {
      seen.add(id)
      files.push({ id, name })
    }
  }
  return files
}

/**
 * List a public Google Drive folder's PDF files (id + name). Downloads the
 * embeddedfolderview listing (see parseDriveFolderListing) and keeps only
 * entries whose name ends in .pdf. Throws a clear error when the link isn't
 * a Drive folder or the folder isn't publicly shared / has no PDFs.
 */
export async function listDriveFolderPdfs(link: string): Promise<DriveFile[]> {
  const id = driveFolderId(link)
  if (!id) {
    throw new Error("That doesn't look like a Google Drive folder link (expected drive.google.com/drive/folders/…).")
  }
  const html = (await download(`https://drive.google.com/embeddedfolderview?id=${id}#list`)).toString('utf8')
  const pdfs = parseDriveFolderListing(html).filter((f) => f.name.toLowerCase().endsWith('.pdf'))
  if (pdfs.length === 0) {
    throw new Error(
      'No PDF files found in that Drive folder. Make sure it\'s shared as "Anyone with the link can view" and contains PDFs.'
    )
  }
  return pdfs
}

/** Download one Drive file (by id) as raw bytes — small tender PDFs download directly, with no virus-scan confirmation step. */
export function downloadDriveFile(id: string): Promise<Buffer> {
  return download(`https://drive.google.com/uc?export=download&id=${id}`)
}

import { download } from './googleImport'

export interface DriveFile {
  id: string
  name: string
}

export interface DriveEntry extends DriveFile {
  isFolder: boolean
}

/** Subfolders whose name mentions "contractor(s)" are never scanned — they hold contractor paperwork, not tender-evaluation PDFs. */
const SKIP_FOLDER_RE = /contractor/i
// Guard against a pathologically deep or self-referential share tree.
const MAX_DEPTH = 6

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
 * entries — files *and* subfolders. That view
 * (drive.google.com/embeddedfolderview?id=…#list) is the stable, no-API-key
 * way to enumerate a publicly shared folder: each item is a
 * `<div class="flip-entry" id="entry-<ID>">` whose `<div class="flip-entry-title">`
 * holds the name, with an `<a href>` that points at `/file/d/<id>` for a file
 * or `/drive/folders/<id>` (or `folderview`) for a subfolder — which is how
 * folder vs file is told apart here. Deduped by id, in listing order.
 */
export function parseDriveFolderListing(html: string): DriveEntry[] {
  const entries: DriveEntry[] = []
  const seen = new Set<string>()
  // Split on each entry's start so a chunk holds one item's href + title.
  const chunks = html.split(/id="entry-/).slice(1)
  for (const chunk of chunks) {
    const idMatch = /^([a-zA-Z0-9_-]+)/.exec(chunk)
    if (!idMatch) continue
    const id = idMatch[1]
    if (seen.has(id)) continue
    const titleMatch = /flip-entry-title[^>]*>([^<]+)</.exec(chunk)
    if (!titleMatch) continue
    const name = titleMatch[1].trim()
    if (!name) continue
    const hrefMatch = /href="([^"]+)"/.exec(chunk)
    const href = hrefMatch ? hrefMatch[1] : ''
    const isFolder = /\/folders\/|folderview/i.test(href)
    seen.add(id)
    entries.push({ id, name, isFolder })
  }
  return entries
}

/**
 * Walk a Drive folder tree from `rootId`, returning every PDF file found at
 * any depth. Recurses into subfolders — the shared link is often a parent of
 * per-work subfolders holding the PDFs — but never descends into a folder
 * whose name mentions "contractor" (SKIP_FOLDER_RE), and guards against
 * cycles/excessive depth. `fetchFolderHtml` returns a folder id's
 * embeddedfolderview HTML (injected so this is unit-testable without network).
 */
export async function collectFolderPdfs(
  rootId: string,
  fetchFolderHtml: (folderId: string) => Promise<string>
): Promise<DriveFile[]> {
  const pdfs: DriveFile[] = []
  const visited = new Set<string>()

  async function walk(folderId: string, depth: number): Promise<void> {
    if (visited.has(folderId) || depth > MAX_DEPTH) return
    visited.add(folderId)
    const entries = parseDriveFolderListing(await fetchFolderHtml(folderId))
    for (const e of entries) {
      if (e.isFolder) {
        if (SKIP_FOLDER_RE.test(e.name)) continue
        await walk(e.id, depth + 1)
      } else if (e.name.toLowerCase().endsWith('.pdf')) {
        pdfs.push({ id: e.id, name: e.name })
      }
    }
  }

  await walk(rootId, 0)
  return pdfs
}

/**
 * List every PDF in a public Google Drive folder link, recursing through
 * subfolders (skipping any "contractor(s)" folder). Throws a clear error
 * when the link isn't a Drive folder or the (public) tree has no PDFs.
 */
export async function listDriveFolderPdfs(link: string): Promise<DriveFile[]> {
  const id = driveFolderId(link)
  if (!id) {
    throw new Error("That doesn't look like a Google Drive folder link (expected drive.google.com/drive/folders/…).")
  }
  const fetchFolderHtml = async (folderId: string) =>
    (await download(`https://drive.google.com/embeddedfolderview?id=${folderId}#list`)).toString('utf8')
  const pdfs = await collectFolderPdfs(id, fetchFolderHtml)
  if (pdfs.length === 0) {
    throw new Error(
      'No PDF files found in that Drive folder (or its subfolders). Make sure it\'s shared as "Anyone with the link can view" and contains PDFs.'
    )
  }
  return pdfs
}

/** Download one Drive file (by id) as raw bytes — small tender PDFs download directly, with no virus-scan confirmation step. */
export function downloadDriveFile(id: string): Promise<Buffer> {
  return download(`https://drive.google.com/uc?export=download&id=${id}`)
}

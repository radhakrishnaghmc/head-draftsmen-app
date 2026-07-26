import { download } from './googleImport'

export interface DriveFile {
  id: string
  name: string
}

export interface DriveEntry extends DriveFile {
  isFolder: boolean
}

// Google's own "<name>_files" asset folders (saved-webpage sidecars) never
// hold tender PDFs — not descended into.
const ASSET_FOLDER_RE = /_files$/i
// Guard against a pathologically deep or self-referential share tree.
const MAX_DEPTH = 8

/**
 * The e-procurement portal's tender-evaluation documents have consistent
 * filenames — the commercial "L1" selection page and the "Stage Selected
 * Form" responsiveness/evaluation pages. A bidder's own uploaded documents
 * (PAN, Aadhaar, GST, IT returns, registration, no-blacklist, …) never do.
 * That filename is the one reliable way to tell them apart, since the real
 * share tree nests them together inconsistently — sometimes the tender PDFs
 * sit beside an agency subfolder, sometimes inside that agency's "Common
 * Documents" folder alongside the bidder docs — so folder-name/structure
 * skipping alone can't separate them.
 */
const TENDER_PDF_RE = /stage\s*selected\s*form|^l\s*-?\s*1\b/i

function isTenderPdf(name: string): boolean {
  const n = name.trim()
  return n.toLowerCase().endsWith('.pdf') && TENDER_PDF_RE.test(n)
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
 * Walk a Drive folder tree from `rootId`, returning every tender-evaluation
 * PDF (by filename — see isTenderPdf / TENDER_PDF_RE) found at any depth. The
 * whole tree is traversed because the real share layout nests the tender
 * PDFs inconsistently (sometimes in the per-work folder, sometimes inside an
 * agency's "Common Documents" subfolder), so the filename filter — not the
 * folder structure — is what keeps a bidder's own documents (PAN/GST/…) out.
 * "<name>_files" asset folders are skipped, and cycles/excessive depth are
 * guarded. `fetchFolderHtml` returns a folder id's embeddedfolderview HTML
 * (injected so this is unit-testable offline).
 */
export async function collectFolderPdfs(
  rootId: string,
  fetchFolderHtml: (folderId: string) => Promise<string>,
  concurrency = 8
): Promise<DriveFile[]> {
  const pdfs: DriveFile[] = []
  const visited = new Set<string>([rootId])
  const queue: { id: string; depth: number }[] = [{ id: rootId, depth: 0 }]

  // A bounded worker pool over a growing queue: the real tree fans out into
  // 100+ folders (per-NIT → per-work → per-agency → Common Documents), and
  // listing them one network request at a time is far too slow — so up to
  // `concurrency` folders are listed at once, each enqueueing its own
  // subfolders as it resolves. A folder that fails to list is skipped rather
  // than aborting the whole scan.
  let active = 0
  return await new Promise<DriveFile[]>((resolve) => {
    const pump = () => {
      if (queue.length === 0 && active === 0) {
        resolve(pdfs)
        return
      }
      while (active < concurrency && queue.length > 0) {
        const { id, depth } = queue.shift()!
        active++
        fetchFolderHtml(id)
          .then((html) => {
            for (const e of parseDriveFolderListing(html)) {
              if (e.isFolder) {
                if (!ASSET_FOLDER_RE.test(e.name) && depth < MAX_DEPTH && !visited.has(e.id)) {
                  visited.add(e.id)
                  queue.push({ id: e.id, depth: depth + 1 })
                }
              } else if (isTenderPdf(e.name)) {
                pdfs.push({ id: e.id, name: e.name })
              }
            }
          })
          .catch(() => {})
          .finally(() => {
            active--
            pump()
          })
      }
    }
    pump()
  })
}

/**
 * List every tender-evaluation PDF (L1 / Stage Selected Form pages) in a
 * public Google Drive folder link, recursing the whole tree and keeping only
 * those by filename — so a bidder's own uploaded documents are ignored (see
 * collectFolderPdfs). Throws a clear error when the link isn't a Drive folder
 * or the (public) tree has no such PDFs.
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
      'No tender-evaluation PDFs (L1 / Stage Selected Form) found in that Drive folder or its subfolders. Make sure it\'s shared as "Anyone with the link can view".'
    )
  }
  return pdfs
}

/** Download one Drive file (by id) as raw bytes — small tender PDFs download directly, with no virus-scan confirmation step. */
export function downloadDriveFile(id: string): Promise<Buffer> {
  return download(`https://drive.google.com/uc?export=download&id=${id}`)
}

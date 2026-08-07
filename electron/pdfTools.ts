import { PDFDocument } from 'pdf-lib'
import * as fs from 'fs'
import * as path from 'path'

/** How many pages a PDF has (encrypted-but-openable PDFs are read too). */
export async function pdfPageCount(srcPath: string): Promise<number> {
  const doc = await PDFDocument.load(fs.readFileSync(srcPath), { ignoreEncryption: true })
  return doc.getPageCount()
}

/** Append every page of each source PDF, in the given order, into one file. */
export async function mergePdfFiles(srcPaths: string[], outPath: string): Promise<void> {
  const merged = await PDFDocument.create()
  for (const p of srcPaths) {
    const src = await PDFDocument.load(fs.readFileSync(p), { ignoreEncryption: true })
    const pages = await merged.copyPages(src, src.getPageIndices())
    for (const pg of pages) merged.addPage(pg)
  }
  fs.writeFileSync(outPath, await merged.save())
}

/** A 1-based inclusive page range. */
export type PageRange = [number, number]

/**
 * Split `srcPath` into one PDF per range (each written into `outDir`). When
 * `ranges` is null, every page becomes its own file. Ranges are clamped to the
 * document and empty ones skipped. File names carry the page label, e.g.
 * "Report p3.pdf" or "Report p4-6.pdf".
 */
export async function splitPdfFile(srcPath: string, outDir: string, ranges: PageRange[] | null): Promise<string[]> {
  const src = await PDFDocument.load(fs.readFileSync(srcPath), { ignoreEncryption: true })
  const total = src.getPageCount()
  const base = path.basename(srcPath, path.extname(srcPath))
  const groups: PageRange[] = ranges ?? Array.from({ length: total }, (_, i) => [i + 1, i + 1] as PageRange)
  const files: string[] = []
  for (const [start, end] of groups) {
    const indices: number[] = []
    for (let p = Math.max(1, start); p <= Math.min(total, end); p++) indices.push(p - 1)
    if (indices.length === 0) continue
    const out = await PDFDocument.create()
    const pages = await out.copyPages(src, indices)
    for (const pg of pages) out.addPage(pg)
    const first = indices[0] + 1
    const last = indices[indices.length - 1] + 1
    const label = first === last ? `p${first}` : `p${first}-${last}`
    const file = path.join(outDir, `${base} ${label}.pdf`)
    fs.writeFileSync(file, await out.save())
    files.push(file)
  }
  return files
}

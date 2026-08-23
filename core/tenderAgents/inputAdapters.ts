/**
 * Every agent in this folder reads exactly one shape: `lines: string[]`, in
 * top-to-bottom reading order, one entry per visually distinct line/row. The
 * detectors themselves never know or care where those lines came from — a
 * digital PDF's own text layer, a photographed page run through this app's
 * local OCR engine, or a spreadsheet's rows all reduce to this one shape, so
 * the exact same 11 detectors run unchanged against any of them.
 *
 * This file holds that normalization step for the non-PDF sources — the
 * PDF path already exists and needs no new adapter: pdf.js's own text
 * reconstruction (src/pdfToText.ts, the same technique this file's tests
 * replicate for standalone use) already produces this exact shape, and every
 * existing agent test uses lines pulled that way from real PDFs.
 */

/** One OCR-detected line of text and its vertical position on the page.
 * Matches electron/ocr.ts's OcrLine shape structurally (duck-typed, not
 * imported — core/ stays free of any main-process-only import so it can be
 * used from the renderer too). */
export interface OcrLikeLine {
  text: string
  top: number
}

/**
 * A photo or scanned PDF page, after this app's local OCR engine has read
 * it: put its detected lines back into top-to-bottom reading order. Mirrors
 * what electron/main.ts's ocrEstimatePhotos handler already does inline for
 * the Estimate-from-Photos feature — pulled out here as its own tested,
 * reusable step so any agent can run against an OCR'd photo (a phone photo
 * of a printed L1 sheet or Intimation letter, not just a digitally exported
 * PDF of the same page) the same way it runs against a digital PDF's text.
 * Blank/whitespace-only detections are dropped, matching every agent's
 * existing assumption that a "line" always has real content.
 */
export function linesFromOcr(ocrLines: readonly OcrLikeLine[]): string[] {
  return [...ocrLines]
    .sort((a, b) => a.top - b.top)
    .map((l) => l.text)
    .filter((t) => t.trim().length > 0)
}

/**
 * A spreadsheet's rows (e.g. a manually compiled tender comparison workbook,
 * or a Data Sheet carrying the same fields): flatten each row's cells into
 * one line, in row order, skipping empty cells. A row becomes one "line" the
 * same way one printed line of a PDF does — enough for the label-anchored
 * regexes every agent uses, since a label and its value sitting in adjacent
 * cells of the same row collapse to "<Label> <value>", exactly how a PDF
 * line reads. Column layout beyond that isn't reconstructed: a value that
 * spans or wraps across several rows won't rejoin the way a wrapped PDF
 * cell does (see priceBidRow.ts) — only genuinely one-row-per-value sheets
 * read cleanly this way.
 */
export function linesFromExcelRows(
  rows: ReadonlyArray<ReadonlyArray<string | number | null | undefined>>
): string[] {
  return rows
    .map((row) =>
      row
        .filter((c) => c != null && String(c).trim().length > 0)
        .map((c) => String(c).trim())
        .join(' ')
    )
    .filter((line) => line.length > 0)
}

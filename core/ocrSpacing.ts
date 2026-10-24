// The PaddleOCR recognizer sometimes drops the spaces between words at one
// particular render resolution of a portal page (real report: the View
// Bidders "Name of Work" came back as "Constrcutionof Mahila bhavan over
// existing seniorcitizenbuilding inTSllCcolonyparkinShapur…" at the app's
// default scale, yet read perfectly at a slightly lower or higher one). The
// glitch is resolution-specific, not page-specific, so the fix is to re-read
// at another scale — this scores a result so the caller can tell.

/** Longest run of letters/digits with no space, treated as "words glued". */
const GLUED_RUN = 20

/** Number of OCR lines containing a suspiciously long unbroken word. */
export function gluedLineCount(lines: string[]): number {
  let n = 0
  for (const line of lines) {
    for (const token of line.split(/[\s,;:()\-/.]+/)) {
      if (token.length >= GLUED_RUN) {
        n++
        break
      }
    }
  }
  return n
}

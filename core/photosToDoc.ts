import * as XLSX from 'xlsx'

/**
 * Convert OCR'd text (photos / a scanned PDF, one recognised line per input
 * line) into either a Word (.docx) or an Excel (.xlsx) document. The renderer
 * shows this text in an editable box first, so whatever the user hands here is
 * already reviewed — these functions just format it, they don't re-interpret it.
 */

const esc = (s: string): string =>
  String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')

/**
 * Split the reviewed text into lines. Blank lines are kept (they carry the
 * page/paragraph spacing the user sees in the preview); trailing whitespace on
 * each line is trimmed.
 */
export function splitLines(text: string): string[] {
  return text.replace(/\r\n?/g, '\n').split('\n').map((l) => l.replace(/\s+$/, ''))
}

/**
 * Body HTML for the Word export: one paragraph per line, blank lines becoming
 * empty paragraphs so the gaps between blocks survive. Meant to be handed to
 * core/htmlToDocx's convertHtmlToDocx (which wraps it in a full HTML document).
 */
export function textToParagraphsHtml(text: string): string {
  const paras = splitLines(text).map((line) => (line.trim() === '' ? '<p>&nbsp;</p>' : `<p>${esc(line)}</p>`))
  // An entirely empty input still needs one paragraph so the .docx is valid.
  return paras.length > 0 ? paras.join('') : '<p>&nbsp;</p>'
}

/**
 * Rows for the Excel export: one row per non-empty line, each split into cells
 * on runs of two-or-more spaces so a photographed table spreads across columns
 * while prose stays in the first column. Blank lines are dropped (an empty
 * spreadsheet row carries no meaning).
 */
export function textToRows(text: string): string[][] {
  return splitLines(text)
    .filter((line) => line.trim() !== '')
    .map((line) => line.split(/\s{2,}/).map((c) => c.trim()))
}

/**
 * Build an .xlsx buffer from the reviewed text — a plain grid (no header row is
 * invented; the OCR'd lines are written as-is), one line per row. Ragged rows
 * are fine; Excel simply leaves the short rows' trailing cells empty.
 */
export function buildPhotosWorkbook(text: string, sheetName = 'OCR'): Buffer {
  return buildWorkbookFromRows(textToRows(text), sheetName)
}

/** Build an .xlsx buffer from a 2-D grid of cell strings (the OCR table reconstruction's rows). */
export function buildWorkbookFromRows(rows: string[][], sheetName = 'OCR'): Buffer {
  const aoa = rows.length > 0 ? rows : [['']]
  const sheet = XLSX.utils.aoa_to_sheet(aoa)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, sheet, (sheetName || 'OCR').slice(0, 31))
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) as Buffer
}

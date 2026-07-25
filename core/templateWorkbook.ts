import type * as ExcelJS from 'exceljs'

/**
 * Strips every data-validation rule from a freshly-loaded template
 * worksheet, before any further editing or re-saving.
 *
 * ExcelJS's reader expands a validation's sqref into one model entry *per
 * cell* it covers — a rule meant to cover "however many rows get typed in
 * later" (e.g. `sqref="A7:A65536"`, a common real-world department template
 * pattern) expands to millions of entries. Its writer's own logic for
 * re-coalescing that model back into `<dataValidation>` ranges is buggy at
 * this scale: it emits overlapping/duplicated ranges for what was
 * originally one clean rule — confirmed by round-tripping this app's own
 * bundled boq-template.xlsx and deviation-template.xlsx through ExcelJS
 * with zero edits and seeing the same corruption appear both times. Real
 * Excel refuses to open the result without a "we found a problem with some
 * content" repair prompt.
 *
 * Data validation is an input-time guardrail anyway — meaningless on a
 * document this app is handing over already filled in — so dropping it here
 * avoids the whole bug class rather than working around ExcelJS's specific
 * miscount.
 */
export function stripDataValidations(worksheet: ExcelJS.Worksheet): void {
  // Not in ExcelJS's own .d.ts (it exists at runtime — see
  // node_modules/exceljs/lib/doc/worksheet.js — the type just isn't
  // declared), same gap as the `templateBuffer as unknown as ArrayBuffer`
  // casts this app's other template-filling code already uses for ExcelJS.
  ;(worksheet as unknown as { dataValidations: { model: Record<string, unknown> } }).dataValidations.model = {}
}

/**
 * Deletes every row/column beyond `lastRow`/`lastCol` — undoing a common
 * real-world spreadsheet-authoring habit (selecting/formatting whole rows or
 * columns at once) that otherwise leaves thousands of empty-but-styled cells
 * riding along in the file. Found the same way as stripDataValidations: this
 * app's own bundled boq-template.xlsx reports actualRowCount 14 /
 * actualColumnCount 8 despite its own rowCount/columnCount claiming 1034/
 * 254 — roughly a quarter million phantom styled cells (row 500, column
 * 200 has no value but does have a style) that bloat the worksheet XML to
 * several megabytes uncompressed. That's most of what makes the exported
 * file large and is a plausible cause of the "unreadable content"/failed-
 * to-open behavior reported on an older Excel version, which is far less
 * tolerant of an oversized used range than modern Excel.
 *
 * Rows are truncated by directly shortening the worksheet's internal
 * `_rows` array rather than `worksheet.spliceRows()` — that method has an
 * off-by-one bug for a range reaching exactly to the array's end (the
 * common case here, trimming "everything after the real content"): it
 * computes `nKeep = start + count` and only clears cells up to `nKeep`, but
 * never runs at all once `nKeep` exceeds the array's current length, which
 * it always does for a to-the-end removal — verified directly (rowCount
 * stayed unchanged before and after calling spliceRows this way). Same
 * `as unknown as` gap as stripDataValidations: `_rows` isn't in ExcelJS's
 * own .d.ts, but is a plain array at runtime.
 *
 * Columns don't have this problem — `spliceColumns()`'s per-row `Row.splice`
 * correctly blanks out each cell's value and style for the removed range
 * even when it reaches the row's own last cell (confirmed by inspecting the
 * written file's XML: no more `<c>` entries past `lastCol`) — it just
 * doesn't update the `columnCount`/`cellCount` getters to reflect it, which
 * doesn't matter here since nothing downstream reads those again.
 */
export function trimToContent(worksheet: ExcelJS.Worksheet, lastRow: number, lastCol: number): void {
  if (worksheet.rowCount > lastRow) {
    ;(worksheet as unknown as { _rows: unknown[] })._rows.length = lastRow
  }
  if (worksheet.columnCount > lastCol) {
    worksheet.spliceColumns(lastCol + 1, worksheet.columnCount - lastCol)
  }
}

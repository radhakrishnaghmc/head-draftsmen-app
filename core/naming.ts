/**
 * Build an output file name from a pattern, resolving tokens:
 *   {Template}  → template name (without .docx)
 *   {RowIndex}  → 1-based row number
 *   {ColumnName}→ the value of that data column
 * The result is sanitized for filesystem safety. Pure (no node deps) so it can
 * be shared by the main process and the renderer (live preview).
 */
export function buildFileName(
  pattern: string,
  templateName: string,
  rowIndex: number,
  row: Record<string, string>
): string {
  const templateBase = templateName.replace(/\.docx$/i, '')
  let name = pattern
    .replace(/\{Template\}/g, templateBase)
    .replace(/\{RowIndex\}/g, String(rowIndex + 1))
  name = name.replace(/\{([^{}]+)\}/g, (_, col: string) => row[col] ?? '')
  name = name.replace(/[\\/:*?"<>|]/g, '_').trim()
  return name.length > 0 ? name : `${templateBase}_${rowIndex + 1}`
}

/** Default naming pattern used when the user doesn't provide one. */
export const DEFAULT_NAME_PATTERN = '{Template}_{RowIndex}'

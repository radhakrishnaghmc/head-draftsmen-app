// Escapes a label for use inside a RegExp, then loosens its internal
// whitespace to `\s*` so OCR's occasional dropped/doubled spaces ("Name  of
// the Work" or "Nameofthe Work") still match the same label.
function labelPattern(label: string): string {
  const escaped = label.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  return escaped.replace(/\s+/g, '\\s*')
}

/**
 * Finds a "Label: value" style line among a photo's OCR'd text lines (see
 * electron/ocr.ts's recognizeImage, one array entry per detected line, in
 * reading order) and returns the value — used to fill an intimation
 * template's {{placeholders}} from whatever's printed/handwritten on the
 * uploaded site photos (a signboard, a stamped note, ...) rather than
 * requiring the user to retype it.
 *
 * Handles both layouts a photo commonly has: the value right after the
 * label on the same line ("Date: 12.07.2026"), and the label alone on its
 * own line with the value printed on the next line down. Returns undefined
 * when no line matches the label at all — never a guess.
 */
export function extractLabeledLine(lines: string[], label: string): string | undefined {
  // Word-bounded on both ends so a label like "Date" doesn't match inside
  // "Datebook" or "Updated" — only a line that mentions the label as its
  // own word(s) counts.
  const re = new RegExp(`\\b${labelPattern(label)}\\b\\s*[:\\-–—]?\\s*(.*)$`, 'i')
  for (let i = 0; i < lines.length; i++) {
    const line = (lines[i] ?? '').trim()
    if (!line) continue
    const m = re.exec(line)
    if (!m) continue
    const inline = m[1]?.trim().replace(/^["'“]+|["'”]+$/g, '')
    if (inline) return inline
    for (let j = i + 1; j < lines.length; j++) {
      const next = (lines[j] ?? '').trim()
      if (next) return next
    }
    return undefined
  }
  return undefined
}

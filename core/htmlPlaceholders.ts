const PLACEHOLDER_RE = /\{\{\s*([^{}]+?)\s*\}\}/g

/** Every distinct {{Label}} found in an HTML template, in first-appearance order. */
export function findHtmlPlaceholders(html: string): string[] {
  const seen = new Set<string>()
  const labels: string[] = []
  for (const m of html.matchAll(PLACEHOLDER_RE)) {
    const label = m[1].trim()
    if (label && !seen.has(label)) {
      seen.add(label)
      labels.push(label)
    }
  }
  return labels
}

/**
 * Replaces every {{Label}} occurrence with its resolved value (blank if
 * `values` has nothing for that label) — a plain text swap since the
 * template is HTML, not a zipped .docx that needs XML-aware editing.
 */
export function fillHtmlPlaceholders(html: string, values: Record<string, string>): string {
  return html.replace(PLACEHOLDER_RE, (_match, rawLabel: string) => values[rawLabel.trim()] ?? '')
}

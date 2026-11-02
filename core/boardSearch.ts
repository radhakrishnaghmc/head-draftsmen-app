/**
 * The search box on the Issue Notices reminders and on the Forwarding / QC letter boards: every word typed must appear
 * somewhere in the work name or the agency name (case-insensitive, any order). A blank query matches everything.
 */
export function matchesBoardSearch(query: string, workName?: string, agency?: string): boolean {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean)
  if (words.length === 0) return true
  const hay = `${workName ?? ''} ${agency ?? ''}`.toLowerCase().replace(/\s+/g, ' ')
  return words.every((w) => hay.includes(w))
}

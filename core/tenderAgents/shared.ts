/**
 * Small pieces genuinely shared by more than one detector below — trivial
 * plumbing (number parsing, line-walking), not extraction logic. Each
 * detector's own field-reading logic stays in its own file; nothing about
 * *what* a field means or *how* to recognize it lives here.
 */

/** Parses a comma-formatted rupee amount ("1,23,456.78") into a number, or undefined if it isn't one. */
export function toNumber(s: string): number | undefined {
  const n = Number(s.replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : undefined
}

/** The nearest non-empty trimmed line before index i, or '' if none. */
export function prevNonEmptyLine(lines: string[], i: number): string {
  for (let j = i - 1; j >= 0; j--) {
    const t = lines[j].trim()
    if (t) return t
  }
  return ''
}

/**
 * Whether `line` looks like the tail/head of a wrapped bidder name (e.g.
 * "Contractor" continuing "Kummary Renuka Devi Civil") rather than a number
 * row or one of the page's own control / header rows — shared across every
 * portal page that lists bidders and wraps a long company name across
 * lines: the L1 Commercial-Evaluation price-bid table (L1 Agency, Name of
 * the Agencies) AND the "View Bidders" supplier list (core/viewBiddersPdf.ts,
 * used by the Evaluation Sheet tool) — same reassembly problem, same fix,
 * one shared implementation so it can't drift into two different bugs.
 */
export function isNameContinuation(line: string): boolean {
  const t = line.trim()
  if (!t || t.length > 40 || /\d/.test(t)) return false
  // Reject the page's own control rows and header remnants ("Value",
  // "Amount", "Rank", the "( INR)"/"INR)" currency-unit tag that wraps onto
  // its own line right above the L-1 row on some layouts, "Supplier"/"Edit
  // Bid" on the View Bidders page…) that land beside the bidder rows — they
  // are not part of any agency name. Strip leading/trailing punctuation
  // first: a punctuation-wrapped remnant like "( INR) INR)" would otherwise
  // dodge the whole blacklist below just by starting with "(" instead of
  // one of the listed words — the same gap that once let an "ITEM n /
  // Dated" tag punctuated differently than expected slip past a similarly
  // word-anchored check (see nameOfWork.ts).
  const stripped = t.replace(/^[^a-z0-9]+|[^a-z0-9]+$/gi, '')
  if (!stripped) return false
  return !/^(back|save|reject|icons|view documents|dashboard|company name|department|refund|information technology|https?:|value|amount|rank|select|percentage|estimated|close|server|price bid|inr|bulk download|action|supplier|edit bid|welcome|tender creation|request a callback)\b/i.test(
    stripped
  )
}

/** The price-bid table's numeric row core: "<ECV> <Less|Excess> <pct> <amount> L-<rank>" — every field the row itself carries, in one anchor both priceBidRow.ts (L-1 only) and agencyNames.ts (every rank) match against. */
export const PRICE_ROW = /([\d,]+\.\d{2})\s+(Less|Excess)\s+([\d.]+)\s+([\d,]+\.\d{2})\s+L-?\s*(\d+)\b/i

/** A page's text lines joined into one string with runs of whitespace
 * collapsed — the shape most detectors below scan for a value that a
 * label/tail regex can span even when pdf.js split it across lines. */
export function joinLines(lines: string[]): string {
  return lines.join(' ').replace(/\s+/g, ' ').trim()
}

/**
 * Removes a "(Item No.N)" tag from the start or end of a work name — the tag
 * is pulled out into its own Item No field (see core/bidDocument.ts's
 * extractItemNo), so it shouldn't also print as part of the work name
 * itself. Lives here (not in bidDocument.ts, which re-exports it for every
 * existing caller) because nameOfWork.ts's stripDecorativeWorkNameTags needs
 * it too, and bidDocument.ts itself depends on worksAmounts.ts — importing
 * this from bidDocument.ts into nameOfWork.ts would cycle back through
 * worksAmounts.ts's own matchers, which also need nameOfWork.ts's
 * normalizeWorkNameForMatch.
 */
export function stripItemNoTag(name: string): string {
  return name
    .replace(/^\s*\(item\s*no\.?\s*\d+\)\s*/i, '')
    .replace(/\s*\(item\s*no\.?\s*\d+\)\s*$/i, '')
    .trim()
}

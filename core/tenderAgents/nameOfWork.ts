import { joinLines } from './shared'

// The field labels that follow "Name of Work" on the page — reaching one ends
// the work-name value. "Works Percentage" is the merged value of the next two
// fields (Tender Category = "Works", Tender Evaluation Type = "Percentage")
// that this page's layout drops onto one line right after the work name.
// The NIT's own "ITEM <n> Dated:<date>" tag — the wrapped continuation of the
// "Enquiry/IFB/Tender Notice Number" cell — routinely lands on its own line
// right next to (and, depending on the PDF's exact y-rounding, sometimes
// immediately above) the Name of Work value, and would otherwise be swept up
// as if it were part of the title (e.g. "ITEM 5 Dated:12.08.2026 Laying of CC
// Road …").
//
// This exact tag has already needed re-fixing multiple times — every office/
// circle's own PDF export punctuates the gap between the item number and
// "Dated"/"Dt" a little differently ("ITEM 5 Dated:…", "ITEM 4 ,Dated:…", and
// whatever the next one turns out to be). Patching in one more literal
// punctuation mark each time a new sample shows up is exactly how this bug
// keeps coming back — so instead of an exact separator, accept ANY run of
// non-alphanumeric characters (spaces, commas, colons, dashes, parens, …)
// between the number and the label, which absorbs punctuation variants this
// code has never actually seen.
const ITEM_DATED_LINE = /^item\s*\d+[^a-z0-9]*(?:dated|dt)\b\.?\s*:?\s*[\d./-]*[^a-z0-9]*$/i

function isWorkNameBoundary(line: string): boolean {
  const t = line.trim()
  if (ITEM_DATED_LINE.test(t)) return true
  // A bare "Works" line, exact-matched (not a prefix — unlike the others
  // below — since a real title could legitimately contain the word
  // "Works" as part of a longer line): one real Serilingampally/Ameenpur
  // office sheet ("ee-ptcu-ghmc") splits "Tender Category Works" onto two
  // separate lines instead of landing "Works" glued onto "Works Percentage"
  // like every other office sheet seen so far, so it slipped past that
  // prefix check and got appended to the title.
  if (/^works$/i.test(t)) return true
  return /^(works\s+percentage|tender\s+category|tender\s+type|tender\s+evaluation\s+type|estimated\s+contract|price\s+bid|bid\s+submission)\b/i.test(
    t
  )
}

// A line that's one of the page's own field labels (not part of the work
// name) — so it's never picked up as the value line preceding "Name of Work".
// "Number" alone is a column-header fragment that lands on its own line right
// above the title on some layouts (between an "(Item No.N)" tag and the title
// itself) — exact-line matched (not a prefix, unlike the others below) since
// it's a single common word that could otherwise false-match real title text.
function isFieldLabelLine(line: string): boolean {
  const t = line.trim()
  return (
    t === '' ||
    /^number$/i.test(t) ||
    /^(notice number|tender id|enquiry|ifb|current tender|commercial evaluation|preliminary responsiveness|name of work)\b/i.test(t)
  )
}

/**
 * Pull the "Name of Work" value out of the reconstructed page lines. Two
 * layouts occur: the label and value on one line ("Name of Work <value> …
 * Tender Category"), and — when the value is long — the value cell wrapping
 * *around* its label, so the label lands on its own line between the two value
 * lines:
 *
 *   Laying of CC Road … (Ward No 276, Pragathi nagar Nizampet   <- value part 1
 *   Name of Work                                                <- label alone
 *   Circle-58, Quthbullapur Zone CMC) (Reserved for SC)         <- value part 2
 *   Works Percentage                                            <- next field
 *
 * A single-line regex captures only the part after the label, dropping the
 * first line — producing a fragment that then mis-matches a different work.
 * So: take any text after the label on its line, else the value line just
 * above it, then append the following lines until the next field label.
 */
function extractFromLabelBlock(lines: string[]): string | undefined {
  const li = lines.findIndex((l) => /name of work/i.test(l))
  if (li < 0) return undefined

  const parts: string[] = []
  const after = lines[li].replace(/^.*?name of work\s*/i, '').trim()
  if (after) {
    parts.push(after)
  } else {
    // Label alone: the value's first part wraps ABOVE the label. Collect EVERY
    // preceding value line (a long title routinely spans two or more), walking
    // up until the previous field label (e.g. "Notice Number") or a boundary —
    // not just the single line immediately above, which used to drop the title
    // and leave only its "…under Municipal General Funds…" tail.
    const above: string[] = []
    for (let j = li - 1; j >= 0; j--) {
      const t = (lines[j] ?? '').trim()
      if (!t) continue
      if (isFieldLabelLine(t) || isWorkNameBoundary(t)) break
      above.push(t)
    }
    above.reverse()
    parts.push(...above)
  }
  for (let j = li + 1; j < lines.length; j++) {
    const t = lines[j].trim()
    if (!t) continue
    if (isWorkNameBoundary(t)) break
    parts.push(t)
    if (parts.join(' ').length > 400) break
  }
  // Belt-and-suspenders: if the ITEM/Dated tag ended up glued onto the same
  // line as the title instead of its own line (pdf.js's y-rounding can go
  // either way), strip it off the front rather than relying solely on the
  // line-boundary check above.
  const name = parts
    .join(' ')
    // Same punctuation-agnostic separator as ITEM_DATED_LINE above — this
    // strip must stay in lockstep with that boundary check, or a variant it
    // now recognizes as the tag would still fail to actually strip here.
    .replace(/^item\s*\d+[^a-z0-9]*(?:dated|dt)\b\.?\s*:?\s*[\d./-]*[^a-z0-9]*/i, '')
    .replace(/\s+/g, ' ')
    .trim()
  return name || undefined
}

/**
 * Agent: Name of Work — the tender's title (e.g. "Laying of CC Road from …
 * ward no 275 in Nizampet Circle-58, Quthbullapur Zone, CMC"), read from the
 * L1 sheet. This is the key later matched against the Works List.
 */
export function detectNameOfWork(lines: string[]): string | undefined {
  const primary = extractFromLabelBlock(lines)
  if (primary) return primary
  // Fallback for a layout where the label+value land on one unbroken line
  // that extractFromLabelBlock's line-array walk didn't anchor on for some
  // reason — scan the whole joined page text instead.
  const joined = joinLines(lines)
  const work = /Name of Work\s+(.+?)\s+(?:Tender Category|Tender Type|Estimated Contract)\b/i.exec(joined)
  return work ? work[1].replace(/\s+/g, ' ').trim() : undefined
}

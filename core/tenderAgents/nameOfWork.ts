import { joinLines, stripItemNoTag } from './shared'

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

// "(Reserved for SC)", "Reserved for ST Only", "reserved for Waddera",
// "reserved for Vaddera", "reserved for WLCCS", "Reserved for Waddera/Sagara"
// (a real compound category seen on a Gajularamaram Circle-57 sheet — two
// community names joined by "/", not just one) — offices tag a reserved
// work's category directly in its own title, right alongside (or instead of)
// a separate Reservation column. A trailing "only"/closing paren is common
// but not captured — the category itself is the run of slash-joined
// word(s)/abbreviation(s) right after "for".
const RESERVED_FOR = /reserved\s+for\s+([A-Za-z]+(?:\s*\/\s*[A-Za-z]+)*)/i

/**
 * Whether the work is marked reserved for ANY category — not just SC/ST.
 * Offices reserve works for SC, ST, Waddera, Vaddera, WLCCS and others; a
 * detector hard-coded to SC/ST (as core/workOrderAgreement.ts's
 * reservationFromRow and core/noteSubmitted.ts's isReservedExempt each used
 * to be, independently and inconsistently) silently failed to exempt EMD for
 * every other category. Reservation exempts EMD only — it does NOT exempt
 * ASD: core/worksAmounts.ts's computeWorkAmounts charges ASD at (Tender
 * Percentage − 25%) × ECV once the quote exceeds 25%, for a reserved work
 * exactly the same as an open one.
 */
export function isReservedWork(workName: string): boolean {
  return RESERVED_FOR.test(workName ?? '')
}

/** The reservation category tag itself ("SC", "ST", "Waddera", "Vaddera", "WLCCS", …) as written in the work name — "" when isReservedWork is false. */
export function reservationCategory(workName: string): string {
  return RESERVED_FOR.exec(workName ?? '')?.[1] ?? ''
}

// A Works List "Reservation" column value that only flags WHETHER a work is
// reserved, without naming the category — core/worksTenderUpdate.ts writes
// exactly "Yes"/"No" there automatically. Not usable as the printed category
// text: a real category is always read from the work name itself instead.
const RESERVATION_FLAG_ONLY = /^(yes|no|none|general|open|nil|n\/?a|-|not\s+reserved|true|false)$/i

/**
 * The best available reservation category for a Works List row: the work
 * name's own "reserved for <category>" tag when present (always a real
 * category, e.g. "SC"/"Waddera"), else an explicit "Reservation" column
 * value — but only when that value actually names a category, not a bare
 * Yes/No flag. Shared by core/workOrderAgreement.ts (Forwarding Slip) and
 * core/noteSubmitted.ts (Note Submitted's EMD/ASD clause), which each used
 * to implement (and disagree on) this independently.
 */
export function reservationCategoryFromRow(row: Record<string, string>): string {
  const fromName = reservationCategory(row['Name of the work'] ?? '')
  if (fromName) return fromName
  const explicit = (row['Reservation'] ?? '').trim()
  return explicit && !RESERVATION_FLAG_ONLY.test(explicit) ? explicit : ''
}

// A trailing "(Reserved for SC)" / "Reserved for Waddera Only" tag — stripped
// wholesale (not just the category capture RESERVED_FOR extracts) so it
// doesn't sit in the way of an identity match against the Works List, which
// very often has the plain work name with no such tag at all.
const TRAILING_RESERVED_FOR = /\(?\s*reserved\s+for\s+[a-z]+(?:\s*\/\s*[a-z]+)*(?:\s+only)?\s*\)?\s*$/i

// "(Recall)", "(1st Recall)", "(1ST RECALL)", "(2nd Call)" — a work re-tendered
// because no agency (or no responsive agency) participated the first time
// round gets this tag appended to its title on the L1 sheet; real samples
// (Gajularamaram Circle-57, Nizampet Circle-58) show it always parenthesized,
// glued directly onto the preceding text with no space, ordinal optional,
// casing inconsistent ("(1st recall)" vs "(1ST RECALL)"). The Works List's
// own stored name is the ORIGINAL tender's name — it was never re-tagged on
// every recall — so this must be stripped before matching too, the same as
// the reservation tag above.
const TRAILING_RECALL_OR_CALL = /\(?\s*\d*\s*(?:st|nd|rd|th)?\s*(?:re-?call|call)\s*\)?\s*$/i

/**
 * Strips every decorative tag a work name can carry that has nothing to do
 * with the work's own identity — item number, reservation category, recall/
 * call round — leaving the plain name the Works List itself was very likely
 * entered under. Every one of the app's Works-List-row matchers used to
 * normalize (lowercase, collapse whitespace) WITHOUT this step first, so an
 * L1 sheet's "…CMC (Reserved for SC)" or "…CMC(1ST RECALL)" title matched
 * nothing even when the exact same work sat right there in the Works List
 * under its plain, untagged name — reported as "name of the work is not in
 * the works list" for a work that plainly was. Tags can stack in either
 * order ("…CMC (Reserved for SC) (1st Recall)"), so this strips repeatedly
 * from the end until nothing more comes off, rather than a single pass.
 */
export function stripDecorativeWorkNameTags(name: string): string {
  let s = stripItemNoTag((name ?? '').trim())
  let changed = true
  while (changed) {
    changed = false
    for (const re of [TRAILING_RESERVED_FOR, TRAILING_RECALL_OR_CALL]) {
      const next = s.replace(re, '').trim()
      if (next !== s) {
        s = next
        changed = true
      }
    }
  }
  return s
}

/**
 * The single canonical normalizer for "does this uploaded work name match
 * this Works List row" — strip decorative tags first (see
 * stripDecorativeWorkNameTags's own doc comment for why), then case/
 * whitespace-fold what's left. core/worksTenderUpdate.ts, core/
 * worksAmounts.ts's applyEcvFromBoq, core/scheduleA.ts's findWorksRowByName
 * and core/monitoringImport.ts's mergeMonitoringRows each used to define
 * their own copy of the case/whitespace half of this WITHOUT the tag-
 * stripping half — four independent normalizers, none of which could match
 * a reserved or recalled work's L1-sheet title against the Works List's own
 * plain, untagged entry for that same work.
 */
export function normalizeWorkNameForMatch(s: string | undefined): string {
  return stripDecorativeWorkNameTags(s ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
}

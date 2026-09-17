import type { RateEntry } from './rateDatabase'
import { normalizeUnit } from './rateDatabase'

/**
 * Parses the Telangana Common Schedule of Rates (SoR) PDF's item tables
 * ("Old Material Rates for Electrical items", "Plumbing Items", "CIVIL
 * ITEMS", "Public Health Items", etc.) into the same {@link RateEntry} shape
 * core/rateDatabase.ts produces from a Data Sheet workbook, so both sources
 * feed the same matching pipeline (core/technicalSanction.ts).
 *
 * Doesn't touch the PDF binary itself — takes the reconstructed text lines
 * a renderer-side pdf.js pass produces (see src/pdfToText.ts's
 * pdfPagesToTextLinesFromData, one visual row per line, columns already
 * left-to-right ordered by x-position). That split keeps this module
 * environment-agnostic (runs under both the renderer and the main process,
 * like every other core/ file) and unit-testable with plain string arrays.
 *
 * The SoR's tables share one row shape across every Part of the document:
 * a serial number, an optional lettered/roman sub-variant, a (possibly
 * multi-line-wrapped) description, then a trailing "<qty> <unit> <rate>"
 * column group. Many rows don't carry their own item code — instead a
 * "ELEC- 1.9.1 Supply of -------- Makes : ..." group-heading line (code
 * starts the line, no leading serial number) introduces a run of lettered
 * rows that share its spec and only differ by the trailing rate — mirrors
 * the xlsx sub-variant convention core/rateDatabase.ts's
 * extractBuildupEntries already handles for Data Sheet workbooks.
 */

// A finished row's trailing "<qty> <unit-token> <rate>" column group —
// present on every real item row regardless of which Part/chapter it's in.
// Captures the description text ahead of it, the unit token (normalized via
// normalizeUnit), and the rate.
const UNIT_TOKEN =
  '(?:No|Nos|RM|Mtr|Mtrs|Mts|Metre|Sq\\.?\\s?in\\.?|Sqin|Sqm|Sqmt|Sqmts|Sft|Cum|Kg|Mt|Set|Sets|Job|Each|Pair|Sqmm|Ltr|Litre|Day|Hour)'
const ROW_TAIL_RE = new RegExp(`^(.*?)\\s+(\\d+(?:\\.\\d+)?)\\s*${UNIT_TOKEN}\\s+([\\d,]+(?:\\.\\d+)?)\\s*$`, 'i')

// A row's opening: a bare serial number, optionally followed by a
// lettered/roman sub-variant marker naming which line in the current group
// this row's rate belongs to (e.g. "380 d 6/10A 3 Pin ..."). Requires
// trailing whitespace after the number so a heading like "5.1 CEILING AND
// EXHAUST FANS" (no space between "5" and ".1") never misreads as a serial.
const ROW_START_RE = /^(\d{1,5})\s+(?:(\(?[a-z]\)?|[ivx]{1,4})\s+)?(.*)$/i

// A row's trailing "<qty> <unit> <rate>" column group, standalone with no
// description ahead of it — the shape a continuation line takes when a
// row's description column wraps onto its own line(s) and the qty/unit/rate
// column lands on a *following* line by itself. Indistinguishable from a
// bare "<serial>" row-start by ROW_START_RE alone (both open with a plain
// number), so this is checked first, while a row is already open, before
// ROW_START_RE gets a chance to misread it as the next item.
const BARE_TAIL_RE = new RegExp(`^\\d+(?:\\.\\d+)?\\s*${UNIT_TOKEN}\\s+[\\d,]+(?:\\.\\d+)?\\s*$`, 'i')

// A departmental item code opening a group-heading line or a self-contained
// row, e.g. "ELEC- 1.9.1", "TBSP-G.I-01", "TBSC-S.I-10". Letters (with an
// optional embedded space/hyphen after the department prefix), hyphenated
// segments, ending in a digit — anchored to the very start of the line so a
// serial-numbered row is never mistaken for one.
const CODE_RE = /^([A-Za-z]{2,6}-\s*[A-Za-z0-9.\-]*\d)\b\s*(.*)$/

// Page furniture repeated on every page — running headers/footers, the
// column-header row, chapter/part banners — filtered out before the row
// state machine sees any line, so it's never mistaken for heading or
// row-continuation text.
const NOISE_RE =
  /^(Page\s+\d+\s+of\s+\d+|Sl\.?\s*No\.?\s*(Item\s*code|ITEM\s*CODE)?\s*Description|SoR\s+2026-27|Buildings\s*\(Part|CIVIL\s+ITEMS|Old\s+Material\s+Rates|Plumbing\s+Items|Public\s+Health\s+Items|Index\b|INDEX\s+FOR|Guidelines\s+for)/i

function cleanHeading(h: string): string {
  return h
    .replace(/-{3,}/g, ' ') // placeholder dash run standing in for the row's own variable spec text
    .replace(/\bMakes\s*:.*$/i, '') // brand list adds nothing to match against and varies row to row
    .replace(/\s+/g, ' ')
    .trim()
}

function normVariant(v: string | undefined): string | undefined {
  return v?.replace(/[()]/g, '').toLowerCase()
}

function codeWithVariant(code: string | undefined, variant: string | undefined): string | undefined {
  if (!code) return undefined
  const c = code.replace(/\s+/g, '').replace(/-$/, '')
  return variant ? `${c}(${variant})` : c
}

interface GroupState {
  code?: string
  heading: string[]
}

/** Finalizes the currently buffered row (if its joined text carries a real qty/unit/rate tail) into a RateEntry, folding in the current group heading when the row itself has no code of its own. */
function finalizeRow(bufferLines: string[], letter: string | undefined, group: GroupState, sheetName: string): RateEntry | null {
  const joined = bufferLines.join(' ').replace(/\s+/g, ' ').trim()
  const m = ROW_TAIL_RE.exec(joined)
  if (!m) return null
  const [, descRaw, , rateRaw] = m
  const rate = rateRaw.replace(/,/g, '').trim()
  if (!rate || Number(rate) <= 0) return null

  const unitMatch = /\d+(?:\.\d+)?\s*([A-Za-z.]+)\s+[\d,]+(?:\.\d+)?\s*$/.exec(joined)
  const unit = unitMatch ? normalizeUnit(unitMatch[1]) : undefined

  let desc = descRaw.trim()
  let code = group.code
  // A row can carry its own code instead of inheriting the group's (the
  // civil/plumbing tables' plain numbered items, e.g. "TBSP-G.I-01 S& F 15
  // mm ..." with no preceding lettered group at all).
  const codeMatch = CODE_RE.exec(desc)
  if (codeMatch) {
    code = codeMatch[1]
    desc = codeMatch[2].trim()
  }

  const heading = cleanHeading(group.heading.join(' '))
  const full = heading && heading !== desc ? `${heading}${desc ? ' - ' + desc : ''}` : desc
  if (full.replace(/[^a-zA-Z]/g, '').length < 6) return null

  return {
    description: full,
    rate,
    sheet: sheetName,
    breakdown: [],
    ...(unit ? { unit } : {}),
    ...(codeWithVariant(code, normVariant(letter)) ? { code: codeWithVariant(code, normVariant(letter)) } : {})
  }
}

/**
 * Extracts every scheduled-rate item from one Part/chapter's worth of
 * reconstructed PDF text lines (see module doc above for the expected line
 * shape). Safe to call across the whole document's lines at once — table
 * headers change per chapter but the row grammar doesn't, and NOISE_RE
 * clears each chapter's running title/column-header lines as it goes.
 */
// A legitimate wrapped row/heading (description column wrapping, a brand
// list wrapping) only ever runs a couple of lines. Tables this parser
// doesn't understand (e.g. a 3-line-per-item layout that splits an item
// code's own prefix/suffix across separate physical lines around the
// qty/unit/rate line) instead feed it a long run of lines that never close
// a row — without a cap, `rowBuffer`/`group.heading` swallow that whole run
// and the accumulated junk then taints every subsequent row sharing the
// group, snowballing across the rest of the page. Capping and dropping
// stale state here trades completeness (that table's rows go unextracted)
// for correctness (no corrupted description ever reaches the rate index).
const MAX_ROW_BUFFER_LINES = 6
const MAX_HEADING_EXTRA_LINES = 3

export function extractSorRateEntries(lines: string[], sheetName = 'SoR 2026-27'): RateEntry[] {
  const entries: RateEntry[] = []
  const group: GroupState = { heading: [] }
  let rowBuffer: string[] | null = null
  let rowLetter: string | undefined

  function closeRow() {
    if (!rowBuffer) return
    const entry = finalizeRow(rowBuffer, rowLetter, group, sheetName)
    if (entry) entries.push(entry)
    rowBuffer = null
    rowLetter = undefined
  }

  // Appends to the open row buffer and closes it on a tail match; abandons
  // (drops) the row instead of letting it run away if no tail ever lands.
  function pushRow(line: string) {
    rowBuffer!.push(line)
    if (ROW_TAIL_RE.test(rowBuffer!.join(' ').trim())) {
      closeRow()
    } else if (rowBuffer!.length > MAX_ROW_BUFFER_LINES) {
      rowBuffer = null
      rowLetter = undefined
    }
  }

  for (const raw of lines) {
    const line = String(raw ?? '').trim()
    if (!line || NOISE_RE.test(line)) continue

    if (rowBuffer && BARE_TAIL_RE.test(line)) {
      pushRow(line)
      continue
    }

    const rowStart = ROW_START_RE.exec(line)
    if (rowStart) {
      // A new serial number always ends whatever came before — an unfinished
      // row that never reached a rate tail is dropped (its rate/unit almost
      // certainly wrapped onto a *following* line that a stray page break or
      // footer swallowed; nothing salvageable), and any open heading buffer
      // is committed as the group this row (and following lettered rows)
      // belongs to.
      closeRow()
      rowLetter = normVariant(rowStart[2])
      rowBuffer = []
      pushRow(rowStart[3])
      continue
    }

    if (rowBuffer) {
      // Mid-row continuation — the description wrapped, or the trailing
      // qty/unit/rate landed on its own line (both common: the description
      // column and the rate column are wrapped independently).
      pushRow(line)
      continue
    }

    const codeStart = CODE_RE.exec(line)
    if (codeStart) {
      // A code-led line that already carries its own full rate tail is a
      // self-contained row (no lettered sub-items follow it), not a group
      // heading — hand it straight to the row buffer instead of parking it
      // as a heading no row will ever consume.
      if (ROW_TAIL_RE.test(line)) {
        rowBuffer = []
        pushRow(line)
        continue
      }
      group.code = codeStart[1]
      group.heading = [codeStart[2]]
      continue
    }

    // A heading can wrap onto a following line (a long "Makes : ..." brand
    // list) before the first lettered row starts — but only for a couple of
    // lines; beyond that this isn't a wrapped heading, it's unrelated table
    // content this parser doesn't understand, so the group is dropped
    // rather than let every later row silently inherit polluted text.
    if (group.heading.length > 0) {
      if (group.heading.length <= MAX_HEADING_EXTRA_LINES) {
        group.heading.push(line)
      } else {
        group.code = undefined
        group.heading = []
      }
    }
  }
  closeRow()

  return entries
}

/** Flattens SoR text extracted per-page (see pdfPagesToTextLinesFromData) into one rate index, same convention as core/rateDatabase.ts's buildRateIndex. */
export function buildSorRateIndex(pages: string[][], sheetName = 'SoR 2026-27'): RateEntry[] {
  return extractSorRateEntries(pages.flat(), sheetName)
}

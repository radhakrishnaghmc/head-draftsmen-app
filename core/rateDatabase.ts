export interface RateEntry {
  description: string
  rate: string
  sheet: string
  /** The raw material/labour/machinery breakdown rows that produced this rate — for reproducing as a rate-analysis backup sheet in the generated Technical Sanction. */
  breakdown: string[][]
  /**
   * Measuring unit for this rate, as a lowercased canonical token
   * (cum/sqm/rmt/each/kg/…), when the sheet states it — parsed from the
   * "Rate per <unit>"/"Unit = <unit>" line of a buildup block, or the Unit
   * column of a flat rate list. Used to reject a candidate whose unit is
   * incompatible with the estimate item's own unit. Undefined when the sheet
   * doesn't make the unit available.
   */
  unit?: string
  /**
   * The departmental index code this rate sits under, carrying its sub-variant
   * when the block has one (e.g. "RBR-EECD-8(c)"). Estimates that cite the code
   * — their item description begins "RBR-EECD-8(c): …" — can then be matched to
   * exactly this rate, the single most reliable signal available. Undefined for
   * sheets/rows that carry no code.
   */
  code?: string
}

function isNumericLike(s: string): boolean {
  const t = s.trim().replace(/,/g, '')
  return t !== '' && Number.isFinite(Number(t))
}

// Matches the label cell of a block's final costed line — "Grand Total",
// "Rate per cum = (a+b+c)/10", "Total (Rs)", "Basic rate per Cum", etc. The
// wording varies across sheets, but every variant contains "total" or "rate".
const FINAL_LABEL_RE = /\btotal\b|\brate\b/i

// Some sheets present a handful of concrete-grade variants (M10/M15/M20/
// M25/M30/M40) as a small side table instead of separate "Rate per…" lines —
// the grade token stands in for a label, with the applicable rate as the
// row's last populated number (the columns before it are the intermediate
// cement-quantity/cost workings that produced it).
const GRADE_RE = /^m\d{1,3}$/i

// A departmental index code sitting in a block's lead cell — "RBR-EECD-8",
// "BLD-CSTN-1-1", "RBR-FNDN-4". Letters, then hyphen-separated segments ending
// in a digit (the item number). Anchored so a stray descriptive phrase never
// reads as a code.
const INDEX_CODE_RE = /^[A-Za-z]{2,6}-[A-Za-z0-9.\-]*\d\)?$/

// A lettered/roman sub-variant marker occupying a cell of its own ("a", "(c)",
// "ii", "(A)") — the tell that a block is entering one of several sub-parts
// that share the parent item's index code but each carry their own rate.
const VARIANT_MARKER_RE = /^\(?([a-z]|[ivx]{1,4})\)?$/i

// A buildup block's formula/section/category lines — "Unit = cum", "Taking
// output = 225 cum", "Cost for 300 cum = …", "Rate per cum = …", "a) Labour",
// "b) Machinery", "Note : …" — plus a colon-terminated category sub-label
// ("R&B and Urban Area Works :"). None of these is the item's real
// description, so they must never be taken as one, whether they sit alone on a
// line or beside a sub-variant marker.
const NON_DESCRIPTION_RE =
  /^(unit\b|taking\s+output|cost\s+(for|of)\b|rate\s+(per|for)\b|add\b|note\b|say\b|or$|[a-f]\s*[).]\s*(labour|machinery|material|overhead)|c\s*&\s*d\b|b\s*&\s*c\b|d\s*&\s*e\b)/i

/** Whether a cell's text reads like an item's actual specification, rather than a formula/section/category line that just happens to sit on its own. */
function isRealDescription(text: string): boolean {
  const t = text.trim()
  if (t.length < 8 || !/[a-zA-Z]/.test(t)) return false
  if (/:\s*$/.test(t)) return false
  return !NON_DESCRIPTION_RE.test(t)
}

// Markers that only ever appear in a *buildup* sheet (one where each rate is
// worked out over a block of material/labour/machinery rows), never in a flat
// abstract-style rate list. Their presence routes a sheet to the block parser;
// their absence lets the flat-list parser have first go (see
// extractRateEntriesFromGrid).
const BUILDUP_MARKER_RE = /taking\s+output|rate\s+per\s|grand\s+total/i

// Canonical measuring units, mapping the many spellings/abbreviations the rate
// schedules use onto one token each. Unknown tokens map to undefined (no unit
// recorded), so a stray label never invents a bogus unit that later blocks a
// legitimate match.
const UNIT_ALIASES: Record<string, string> = {
  cum: 'cum',
  cm: 'cum',
  cft: 'cft',
  sqm: 'sqm',
  sqmt: 'sqm',
  sqmts: 'sqm',
  sm: 'sqm',
  sft: 'sft',
  rmt: 'rmt',
  rm: 'rmt',
  rmtr: 'rmt',
  mtr: 'rmt',
  metre: 'rmt',
  each: 'each',
  no: 'each',
  nos: 'each',
  kg: 'kg',
  mt: 'mt',
  tonne: 'mt',
  tonnes: 'mt',
  quintal: 'quintal',
  litre: 'litre',
  ltr: 'litre',
  kl: 'kl',
  day: 'day',
  hour: 'hour'
}

/** Canonicalize a raw unit token ("Sqmt", "Rmt", "cum") to its standard form, or undefined if it isn't a recognized unit. */
export function normalizeUnit(raw: string): string | undefined {
  const t = raw.trim().toLowerCase().replace(/[^a-z]/g, '')
  if (!t) return undefined
  return UNIT_ALIASES[t]
}

// The unit named on a "Rate per <unit>" / "Rate per 1 sqm" line, or a "Unit =
// cum" / "Unit: 1 cum" declaration — the digit (a "per 1 …" multiplier) is
// skipped so the alphabetic unit token is captured.
const RATE_PER_UNIT_RE = /\bper\s*\d*\s*([a-zA-Z]+)/i
const UNIT_DECL_RE = /\bunit\b\s*[:=]\s*\d*\s*([a-zA-Z]+)/i

/** The measuring unit a block row reveals, if any — from its "Rate per <unit>" label or a "Unit = <unit>" declaration. */
function unitFromRow(cells: string[]): string | undefined {
  for (const cell of cells) {
    const decl = UNIT_DECL_RE.exec(cell)
    if (decl) {
      const u = normalizeUnit(decl[1])
      if (u) return u
    }
    if (/\brate\b/i.test(cell)) {
      const per = RATE_PER_UNIT_RE.exec(cell)
      if (per) {
        const u = normalizeUnit(per[1])
        if (u) return u
      }
    }
  }
  return undefined
}

// The rate value sits somewhere to the right of the label on the same row,
// usually the very next populated cell — take the first one, not the last,
// since a stray reference number can sit far to the right (e.g. a Remarks
// column) after several blank cells.
function findRateAfter(cells: string[], labelIdx: number): string | null {
  for (let i = labelIdx + 1; i < cells.length; i++) {
    if (cells[i] !== '' && isNumericLike(cells[i])) return cells[i]
  }
  return null
}

/** The index code (with its sub-variant, when known) that identifies a block's rate — e.g. "RBR-EECD-8" or "RBR-EECD-8(c)". */
function codeWithVariant(code: string | undefined, variant: string | undefined): string | undefined {
  if (!code) return undefined
  return variant ? `${code}(${variant})` : code
}

/** Whether a sheet is a worked-out *buildup* sheet (block per rate) rather than a flat abstract-style rate list. */
function looksLikeBuildupSheet(grid: string[][]): boolean {
  for (const row of grid) {
    for (const cell of row) {
      if (BUILDUP_MARKER_RE.test(String(cell ?? ''))) return true
    }
  }
  return false
}

/**
 * Parse a flat abstract-style rate list — a sheet whose items sit one-per-row
 * as "… | Description | Rate | Unit", the finished rate inline on the same
 * row (the "RBR Abstract" / "ABSTRACT OF DATA RATES" tab is the canonical
 * example). These sheets carry the cleanest, most authoritative descriptions
 * in the whole workbook, but the block parser skips them entirely because
 * there's no trailing "Rate per…"/"Grand Total" label to key off.
 *
 * Only fires when a header row genuinely labels a "Description" column and a
 * "Rate" column. A labour/material *input* list ("Sl.no | Labour | … | Rate",
 * mason @ 680/day) has no "Description" header and is handled separately by
 * extractLabourRateList, which gates out the short raw-input rows. Returns null
 * when no such header exists.
 */
function extractFlatRateList(grid: string[][], sheetName: string): RateEntry[] | null {
  let headerRow = -1
  let descCol = -1
  let rateCol = -1
  let unitCol = -1
  let codeCol = -1
  for (let r = 0; r < grid.length && r < 20; r++) {
    const cells = grid[r].map((c) => String(c ?? '').trim())
    const d = cells.findIndex((c) => /description/i.test(c))
    const rt = cells.findIndex((c) => /^rate\b/i.test(c))
    if (d !== -1 && rt !== -1) {
      headerRow = r
      descCol = d
      rateCol = rt
      unitCol = cells.findIndex((c) => /^(unit|uom|per)\b/i.test(c))
      codeCol = cells.findIndex((c) => /index.?code/i.test(c))
      break
    }
  }
  if (headerRow === -1) return null

  const entries: RateEntry[] = []
  for (let r = headerRow + 1; r < grid.length; r++) {
    const cells = grid[r].map((c) => String(c ?? '').trim())
    const description = cells[descCol] ?? ''
    const rate = cells[rateCol] ?? ''
    // A real item row: a worded description and a positive inline rate. Chapter/
    // section header rows ("A | RBR-EECD (Chapter 3)") have no rate and drop out.
    if (description.length < 8 || !/[a-zA-Z]{3,}/.test(description)) continue
    if (!isNumericLike(rate) || Number(rate.replace(/,/g, '')) <= 0) continue
    // Skip *sub-label* rows whose real description lives in the item rows above
    // them — a case/grading/means variant ("CASE II : With Batching Plant",
    // "For Grading I Material", "R&B and Urban Area Works :"). Taking their own
    // cell as the description would attach the rate to a meaningless fragment;
    // these items are extracted in full (with their proper wording and code)
    // from the corresponding buildup sheet instead.
    if (/:\s*$/.test(description)) continue
    if (/^(case\b|for\s+grading|for\s+every|by\s+manual|by\s+mechanical|r\s*&\s*b\b|using\b)/i.test(description)) continue
    const unit = unitCol !== -1 ? normalizeUnit(cells[unitCol] ?? '') : undefined
    const codeCell = codeCol !== -1 ? cells[codeCol] ?? '' : ''
    const code = INDEX_CODE_RE.test(codeCell) ? codeCell : undefined
    entries.push({ description, rate, sheet: sheetName, breakdown: [cells], unit, code })
  }
  return entries
}

// The longest a description can be while still counting as a bare sub-label or
// raw input name ("I st class Mason", "d) H.D.-20 with 560mm opening") that
// leans on a parent row for its real meaning, rather than a self-contained
// finished-item spec. Above it, a row stands on its own.
const FULL_SPEC_MIN_LEN = 40
// A row's own text is long enough, and worded richly enough, to be a real
// searchable finished-item description in its own right — the floor that keeps
// short raw labour/material names ("Man Mazdoor", "white cement") out of the
// finished-item index even after a parent prefix is considered.
const RATE_LIST_MIN_LEN = 25
// A dependent sub-label opening a row ("d) …", "(iii) …", "a) …") — the tell
// that the row priced below a parent item is one of its variants, so it should
// inherit the parent's specification.
const SUBLABEL_START_RE = /^\(?([a-z]|[ivx]{1,4})\)/i

/**
 * Parse a labour/material rate list — a flat "Sl.No | Item | (ref) | Rate |
 * (Unit)" sheet (the "Labour & material" tab) whose second column is the item
 * name rather than a header-labelled "Description". Most of it is raw input
 * prices (a mason's daily wage, a bag of cement) that must never be matched
 * against an estimate's *work* items, so a length gate keeps those short names
 * out; what survives is the sheet's genuine finished/worked-out items —
 * mechanical trowelling, groove cutting, manhole covers, road-surface cutting —
 * each a real BOQ line with its own rate and unit.
 *
 * A finished item priced as several sub-variants under one heading (a manhole
 * cover "Manufacture as per BIS:12592 …" with H.D.-10 / H.D.-20 / H.D.-35 rate
 * rows beneath) has each variant's short label prefixed with the heading it
 * depends on, so the emitted item keeps the full specification. Returns null
 * when the sheet has no serial+rate header to key off.
 */
function extractLabourRateList(grid: string[][], sheetName: string): RateEntry[] | null {
  let headerRow = -1
  let serialCol = -1
  let descCol = -1
  let rateCol = -1
  let unitCol = -1
  for (let r = 0; r < grid.length && r < 20; r++) {
    const cells = grid[r].map((c) => String(c ?? '').trim())
    const rt = cells.findIndex((c) => /^rate\b/i.test(c))
    const sc = cells.findIndex((c) => /^s[lr]?\.?\s*no\.?$|serial/i.test(c))
    if (rt === -1 || sc === -1) continue
    headerRow = r
    rateCol = rt
    serialCol = sc
    descCol = sc + 1
    unitCol = cells.findIndex((c) => /^(unit|uom|per)\b/i.test(c))
    break
  }
  if (headerRow === -1) return null

  const isFullSpec = (d: string): boolean => d.length >= FULL_SPEC_MIN_LEN
  const entries: RateEntry[] = []
  // The most recent heading a following variant row should inherit — a real
  // spec that either priced nothing of its own or opens a multi-variant group.
  let lastParent: string | undefined
  for (let r = headerRow + 1; r < grid.length; r++) {
    const cells = grid[r].map((c) => String(c ?? '').trim())
    const description = cells[descCol] ?? ''
    const rate = cells[rateCol] ?? ''
    const hasSerial = serialCol !== -1 && /^\d/.test(cells[serialCol] ?? '')
    // Only a serial-numbered full spec opens a group its later variant rows can
    // inherit — a variant row that happens to run long ("H.D.-10 with 560mm
    // clear opening") must not itself become the heading for the variants below.
    if (isFullSpec(description) && hasSerial) lastParent = description
    if (!/[a-zA-Z]{3,}/.test(description)) continue
    if (!isNumericLike(rate) || Number(rate.replace(/,/g, '')) <= 0) continue
    // A short, dependent row (no serial of its own, or opening with a "d)"-style
    // sub-label) borrows the heading it sits under so its rate keeps the full
    // specification; a self-contained full spec stands alone.
    const dependent = !isFullSpec(description) && (!hasSerial || SUBLABEL_START_RE.test(description))
    const effective = dependent && lastParent ? `${lastParent} - ${description}` : description
    // Length gate: keeps raw labour/material input names ("I st class Mason",
    // "white cement") out, while every genuine finished item clears it easily.
    if (effective.length < RATE_LIST_MIN_LEN) continue
    const unit =
      unitCol !== -1 ? normalizeUnit(cells[unitCol] ?? '') : normalizeUnit(cells[rateCol + 1] ?? '')
    entries.push({ description: effective, rate, sheet: sheetName, breakdown: [cells], unit })
  }
  return entries
}

/**
 * Extract {description, rate, unit, code} rates from one *buildup* sheet of a
 * rates database — "Telangana/Andhra Pradesh Standard Data" style workbooks
 * (Building Data, RMC, the RBR-* chapters, etc.). Each item spans a block of
 * rows: a lead row carrying an index code and/or serial number together with
 * the item's description (position varies sheet to sheet, so it's found by
 * picking the longest text cell in that row rather than a fixed column), a
 * material/labour/machinery cost breakdown, and a line further down labelled
 * "Grand Total" / "Rate per <unit> = …" / etc. carrying the final worked-out
 * rate.
 *
 * A new block starts wherever column 0 holds an index code, or column 1 holds
 * a bare serial number — both are blank on every other row in the sheet. An
 * item split into lettered/roman sub-variants (a/b/c, i/ii) that each carry
 * their own full spec and rate is emitted as one entry per sub-variant, each
 * with its own description and its own code(variant) — rather than lumping
 * them all under the parent item's lead title. Sheets that follow neither
 * convention simply yield nothing here; that's an acceptable gap rather than a
 * wrong guess.
 */
function extractBuildupEntries(grid: string[][], sheetName: string): RateEntry[] {
  const entries: RateEntry[] = []
  let pendingDescription: string | null = null
  // Rows seen since the start of the current item (or since the last
  // variant's rate line, for a multi-variant block) — becomes that variant's
  // `breakdown`. Reset after every match so each variant only carries its
  // own specific buildup, not the accumulated history of prior variants.
  let blockRows: string[][] = []
  // The index code / sub-variant / unit currently in scope — carried onto each
  // emitted entry so it can be matched by code and filtered by unit.
  let currentCode: string | undefined
  let currentVariant: string | undefined
  let currentUnit: string | undefined

  for (const raw of grid) {
    const cells = raw.map((c) => String(c ?? '').trim())
    const nonEmpty = cells.filter((c) => c !== '')
    if (nonEmpty.length === 0) continue

    const isItemStart = cells[0] !== '' || /^\d+$/.test(cells[1] ?? '')
    if (isItemStart) {
      if (INDEX_CODE_RE.test(cells[0] ?? '')) currentCode = cells[0]
      currentVariant = cells.slice(1).find((c) => VARIANT_MARKER_RE.test(c))?.replace(/[()]/g, '')
      currentUnit = unitFromRow(cells)
      const longest = nonEmpty.reduce((a, b) => (b.length > a.length ? b : a), '')
      pendingDescription = longest.length >= 8 ? longest : null
      blockRows = pendingDescription ? [cells] : []
      continue
    }

    // A lettered/roman sub-variant within the current item (a cell holding just
    // "a"/"(c)"/"ii"). Record which variant we're in, so a code+variant lookup
    // can pinpoint its rate; and when the same row also carries a full spec of
    // its own — a long description, not a bare "Manual Means" sub-label — adopt
    // it as the description for the rate line(s) that follow, so each variant
    // surfaces with its own wording instead of all inheriting the lead title.
    const variantIdx = cells.findIndex((c) => VARIANT_MARKER_RE.test(c))
    if (variantIdx !== -1 && (pendingDescription || currentCode)) {
      currentVariant = cells[variantIdx].replace(/[()]/g, '')
      const longDesc = nonEmpty.find((c) => c.length >= 25 && isRealDescription(c))
      if (longDesc) {
        pendingDescription = longDesc
        blockRows = [cells]
        continue
      }
    }

    // Some sheets give the item-start row only a short title (e.g.
    // "Excavation for Structures") and put the real, much longer
    // description on its own line right after — a lone text cell with
    // nothing else on the row. Prefer whichever candidate is longer, and —
    // when nothing is pending yet — let it seed pendingDescription outright:
    // a handful of sheets never give their top description an index code or
    // serial number at all (it just sits alone on its own line), and a
    // spurious numeric-looking row above it (e.g. a "1  2  3  4  5" column-
    // index line) can otherwise null out anything pending before this point.
    if (nonEmpty.length === 1 && isRealDescription(nonEmpty[0])) {
      if (!pendingDescription || nonEmpty[0].length > pendingDescription.length) {
        pendingDescription = nonEmpty[0]
      }
    }

    if (!pendingDescription) continue
    blockRows.push(cells)
    const rowUnit = unitFromRow(cells)
    if (rowUnit) currentUnit = rowUnit

    const labelIdx = cells.findIndex((c) => FINAL_LABEL_RE.test(c))
    if (labelIdx !== -1) {
      const rate = findRateAfter(cells, labelIdx)
      // Deliberately don't clear pendingDescription after a match: many
      // items have several total/rate lines under one description (e.g.
      // manual vs. mechanical means × depth range, all sharing the same
      // paragraph) — each one is a genuine, distinct rate for that
      // description, only distinguished by structural sub-labels this
      // parser doesn't track. Keeping the description alive until the next
      // item-start row lets every variant surface as its own entry, which is
      // what makes the multi-rate/ambiguous case detectable downstream
      // instead of silently keeping just the first.
      if (rate) {
        entries.push({
          description: pendingDescription,
          rate,
          sheet: sheetName,
          breakdown: blockRows,
          unit: unitFromRow(cells) ?? currentUnit,
          code: codeWithVariant(currentCode, currentVariant)
        })
      }
      blockRows = []
      continue
    }

    const gradeIdx = cells.findIndex((c) => GRADE_RE.test(c))
    if (gradeIdx !== -1) {
      const numeric = cells.filter((c) => c !== '' && isNumericLike(c))
      const rate = numeric[numeric.length - 1]
      if (rate) {
        const grade = cells[gradeIdx].toUpperCase()
        entries.push({
          description: `${pendingDescription} ${grade}`,
          rate,
          sheet: sheetName,
          breakdown: blockRows,
          unit: currentUnit,
          code: codeWithVariant(currentCode, grade)
        })
      }
      blockRows = []
    }
  }

  return entries
}

/**
 * Extract every {description, rate, unit, code} rate from one sheet of a rates
 * database. A flat rate list (finished rates inline, one per row) is parsed
 * row-by-row — an abstract-style sheet with a "Description" header, or a
 * labour/material list keyed off its "Sl.No … Rate" columns; everything else is
 * treated as a worked-out buildup sheet (a block of cost rows per rate). See
 * extractFlatRateList, extractLabourRateList and extractBuildupEntries.
 */
export function extractRateEntriesFromGrid(grid: string[][], sheetName: string): RateEntry[] {
  if (!looksLikeBuildupSheet(grid)) {
    const flat = extractFlatRateList(grid, sheetName)
    if (flat && flat.length > 0) return flat
    const labour = extractLabourRateList(grid, sheetName)
    if (labour && labour.length > 0) return labour
  }
  return extractBuildupEntries(grid, sheetName)
}

/** Flatten every sheet of a rates database workbook into one searchable list. */
export function buildRateIndex(sheets: { name: string; grid: string[][] }[]): RateEntry[] {
  return sheets.flatMap((s) => extractRateEntriesFromGrid(s.grid, s.name))
}

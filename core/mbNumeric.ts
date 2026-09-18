/**
 * Digit repair and arithmetic reconciliation for MB measurement rows.
 *
 * Handwritten figures photographed in a book are read badly by any general
 * OCR model: "15.00" comes back as "IS.00", "0.30" as "0-30" or "630", and a
 * value followed by the clerk's ditto/unit tick as "28.351". Two things make
 * these recoverable without guessing:
 *
 * 1. Every figure on this form is a plain 2-decimal measurement, so a reading
 *    with a missing or misplaced decimal point has only a handful of possible
 *    intended values.
 * 2. The row carries its own check: Contents = No x L x B x D. Only one of
 *    the candidate readings normally satisfies that identity, so the row's
 *    own arithmetic picks the right one.
 *
 * Corrections are deliberately limited to candidate readings of what OCR
 * actually returned (decimal position, stray trailing mark, letter-shaped
 * digits). A row whose figures are read confidently but whose arithmetic does
 * not agree is NOT rewritten — this is a scrutiny tool, and a genuine
 * arithmetic mistake in the book must survive to be seen.
 */

/** Letter shapes an OCR model commonly returns for a handwritten digit. */
const LETTER_DIGITS: Record<string, string> = {
  O: '0',
  o: '0',
  Q: '0',
  D: '0',
  U: '0',
  I: '1',
  l: '1',
  i: '1',
  '|': '1',
  '!': '1',
  L: '1',
  Z: '2',
  z: '2',
  S: '5',
  s: '5',
  b: '6',
  G: '6',
  T: '7',
  B: '8',
  g: '9',
  q: '9',
  A: '4',
  // The recognition head is trained on Chinese text and returns these glyphs
  // for handwritten digit shapes on this form.
  子: '7',
  了: '7'
}

/** Marks an OCR model returns where the handwritten decimal point was. */
const DECIMAL_MARKS = /[·・.,、;:~˜\-–—_'`"/]|小/g

/**
 * Canonical digits-and-one-dot form of an OCR'd figure: letter shapes mapped
 * back to digits, any decimal-mark shape mapped to ".", surrounding noise
 * (unit words, ditto ticks, stray brackets) dropped.
 */
export function normalizeMbFigure(raw: string): string {
  return mbFigureReadings(raw)[0] ?? ''
}

/**
 * The digit/dot strings a cell's OCR text could be a reading of, most likely
 * first. A figure comes back either in one piece ("15.00") or broken into
 * runs by noise the model invented around the ink ("15 00", "31.05um"), so
 * both the longest single run and all runs joined are offered.
 */
function mbFigureReadings(raw: string): string[] {
  const mapped = raw
    .trim()
    .split('')
    .map((ch) => LETTER_DIGITS[ch] ?? ch)
    .join('')
    .replace(DECIMAL_MARKS, '.')
  const runs = (mapped.match(/[0-9.]+/g) ?? []).map(tidyDots).filter((r) => r !== '')
  if (runs.length === 0) return []
  const longest = runs.reduce((a, b) => (digitCount(b) > digitCount(a) ? b : a))
  const joined = tidyDots(runs.join(''))
  const out = digitCount(longest) >= 3 ? [longest, joined] : [joined, longest]
  return out.filter((v, i) => v !== '' && out.indexOf(v) === i)
}

/** Collapses a run to digits with at most the first dot kept. */
function tidyDots(run: string): string {
  const digits = run.replace(/[^0-9]/g, '')
  if (digits === '') return ''
  const firstDot = run.indexOf('.')
  const digitsBeforeDot = firstDot < 0 ? digits.length : run.slice(0, firstDot).replace(/[^0-9]/g, '').length
  if (digitsBeforeDot === 0 || digitsBeforeDot === digits.length) return digits
  return `${digits.slice(0, digitsBeforeDot)}.${digits.slice(digitsBeforeDot)}`
}

function digitCount(s: string): number {
  return s.replace(/[^0-9]/g, '').length
}

/**
 * How many digits a cell's OCR text really holds, counting letter shapes the
 * model returned for digits ("IS.00" holds 4, not 2).
 */
export function mbMappedDigitCount(raw: string): number {
  return digitCount(
    raw
      .split('')
      .map((ch) => LETTER_DIGITS[ch] ?? ch)
      .join('')
  )
}

/**
 * Plausible intended values for one OCR'd figure, best first. Every figure on
 * this form is written to 2 decimals, so the variants are: the reading as
 * given, the same digits with the decimal point placed 2 from the right (a
 * lost point), and the same again after dropping one trailing digit (the
 * clerk's ditto/unit tick read as a digit).
 */
export function mbFigureCandidates(raw: string): string[] {
  // No real digit anywhere in the cell means this is not a figure at all: it
  // is either description text that overflowed into the column or invention
  // from paper grain. Letter-shaped-digit mapping would otherwise turn
  // "dorprisn" into 0.15 and file it as a measurement.
  if (!/[0-9]/.test(raw)) return []
  const ranked: { value: string; rank: number }[] = []
  const push = (v: string, rank: number): void => {
    const n = Number(v)
    if (!Number.isFinite(n) || n <= 0) return
    const canonical = String(n)
    const existing = ranked.find((r) => r.value === canonical)
    if (existing === undefined) ranked.push({ value: canonical, rank })
    else if (rank < existing.rank) existing.rank = rank
  }

  mbFigureReadings(raw).forEach((reading, readingRank) => {
    const dot = reading.indexOf('.')
    const digits = reading.replace(/\./g, '')
    const base = readingRank * 4
    if (dot >= 0) {
      const decimals = digits.length - dot
      if (decimals === 2) push(reading, base)
      else if (decimals > 2) {
        // The extra trailing character is the clerk's unit/ditto tick read as
        // a digit ("29.704" is 29.70 followed by a tick), which is far more
        // common on these sheets than a genuine 3-decimal figure.
        push(`${digits.slice(0, dot)}.${digits.slice(dot, dot + 2)}`, base)
        push(reading, base + 3)
      } else {
        // One decimal: either a lost trailing zero ("15.0" = 15.00) or a
        // point that landed one place early ("63.0" = 6.30).
        push(reading, base)
        if (digits.length >= 3) push(placeDecimal(digits, 2), base + 1)
      }
    } else {
      push(placeDecimal(digits, 2), base)
      if (digits.length > 3) push(placeDecimal(digits.slice(0, -1), 2), base + 1)
      push(digits, base + 2)
    }
  })

  return ranked.sort((a, b) => a.rank - b.rank).map((r) => r.value)
}

function placeDecimal(digits: string, decimals: number): string {
  if (digits.length <= decimals) return `0.${digits.padStart(decimals, '0')}`
  return `${digits.slice(0, digits.length - decimals)}.${digits.slice(digits.length - decimals)}`
}

/**
 * Best reading of one OCR'd figure cell, or '' when the cell holds no figure.
 *
 * A cell with no real digit at all is treated as noise, not as a figure: OCR
 * models invent plausible-looking text ("pwell", "CeoL") from paper grain in
 * a blank cell, and letter-shaped-digit mapping would otherwise turn that
 * invention into a fabricated measurement.
 */
export function readMbFigure(raw: string): string {
  if (!/[0-9]/.test(raw)) return ''
  return mbFigureCandidates(raw)[0] ?? ''
}

/**
 * An L/B/D cell filled as an average — "6.70+6.50" over "2" — read back into
 * its individual figures. The "+" signs are routinely recognised as digits
 * (4, 1, 7) or dropped, so the digit run is re-split structurally instead:
 * each figure on this form carries 2 decimals (3 digits), so `n` averaged
 * figures are 3n digits, and any extra digits sitting at the group
 * boundaries are the "+" signs misread.
 */
export function splitAverageNumerator(raw: string, count: number): string[] | undefined {
  if (count < 2 || count > 4) return undefined
  const plusSplit = raw
    .split('+')
    .map((part) => normalizeMbFigure(part))
    .filter((part) => part !== '')
  if (plusSplit.length === count && plusSplit.every((p) => p.replace(/\./g, '').length === 3)) {
    // A part whose decimal point was lost ("640" for 6.40) still carries the
    // 3 digits every figure on this form has, so the point goes back 2 from
    // the right.
    return plusSplit.map((p) => String(Number(p.includes('.') ? p : placeDecimal(p, 2))))
  }

  const digits = raw
    .trim()
    .split('')
    .map((ch) => LETTER_DIGITS[ch] ?? ch)
    .join('')
    .replace(/[^0-9]/g, '')
  const wanted = count * 3
  if (digits.length === wanted) return groupsOfThree(digits, count)
  // One extra digit per separator: the "+" signs came back as digits.
  if (digits.length === wanted + (count - 1)) {
    let out = ''
    let i = 0
    for (let g = 0; g < count; g++) {
      out += digits.slice(i, i + 3)
      i += 3 + 1
    }
    return groupsOfThree(out, count)
  }
  return undefined
}

function groupsOfThree(digits: string, count: number): string[] {
  const out: string[] = []
  for (let g = 0; g < count; g++) {
    const part = digits.slice(g * 3, g * 3 + 3)
    out.push(String(Number(`${part[0]}.${part.slice(1)}`)))
  }
  return out
}

/**
 * The count in the "No." column, written as a multiplier such as "1x1" (and
 * read by OCR as "IXI", "(x1", "lX)"). Returns 1 when the cell is blank or
 * unreadable — a blank No. column on this form means a single measurement.
 */
/**
 * Readings for an L/B/D cell, best first: the averaged value when the cell
 * holds an average, plus plain figure readings. A cell written as an average
 * whose denominator OCR missed entirely ("6.70+6.50" with the "2" under the
 * bar lost) still yields the averaged readings here, for the row's own
 * arithmetic to choose between.
 */
export function mbDimensionCandidates(raw: string): string[] {
  const lines = raw
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
  if (lines.length === 0) return []
  const out: string[] = []
  const push = (v: string | undefined): void => {
    if (v !== undefined && v !== '' && !out.includes(v)) out.push(v)
  }

  const denominator = lines.length > 1 ? parseMbCount(lines[lines.length - 1]) : undefined
  const numerator = lines.length > 1 ? lines.slice(0, -1).join('') : lines[0]
  if (denominator !== undefined && denominator >= 2 && denominator <= 4) {
    push(averageOf(numerator, denominator))
  }
  // A cell holding two or three figures is an average, not one long number:
  // prefer the averaged reading over reading the digit run as a figure.
  const looksAveraged = numerator.includes('+') || numerator.replace(/[^0-9]/g, '').length >= 6
  if (looksAveraged) {
    push(averageOf(numerator, 2))
    push(averageOf(numerator, 3))
  }
  push(mbFigureCandidates(lines[0])[0])
  // The denominator is small and often swallowed by the fraction bar, so an
  // unsplit numerator is still offered as a 2- and 3-way average.
  for (const n of [2, 3]) push(averageOf(numerator, n))
  for (const c of mbFigureCandidates(lines[0]).slice(1)) push(c)
  return out
}

/**
 * Repair candidates for a dimension cell: the same figure with a digit swapped
 * for a shape the model confuses it with. Deliberately kept out of
 * mbDimensionCandidates — these are only offered to reconcileMbRow's repair
 * pass, one cell per row, so a guessed digit can never reach the sheet unless
 * the row's own arithmetic singles it out.
 */
export function mbDimensionVariants(raw: string): string[] {
  const out = digitConfusionVariants(mbDimensionCandidates(raw).slice(0, 2))
  // A cell written as an average whose fraction bar and denominator were lost
  // ("6.20+6.50" with no "2" under it) still has to be repaired on its parts,
  // not on the mean — see mbAverageVariants.
  const numerator = raw.split(/\r?\n/)[0] ?? ''
  if (numerator.includes('+') || numerator.replace(/[^0-9]/g, '').length >= 6) {
    for (const n of [2, 3]) {
      const parts = splitAverageNumerator(numerator, n)
      if (parts === undefined) continue
      for (const v of mbAverageVariants(parts, n)) if (!out.includes(v)) out.push(v)
    }
  }
  return out
}

/**
 * Repair candidates for an averaged cell ("6.70+6.50" over 2): the average
 * recomputed with one digit of one part corrected. The average itself is
 * hopeless to repair digit by digit — a wrong digit in a part moves the mean
 * by a fraction the digit shapes never produce — so the repair has to happen
 * on the parts and be averaged again.
 */
export function mbAverageVariants(parts: string[], denominator: number): string[] {
  if (denominator <= 0) return []
  const values = parts.map((p) => Number(p))
  if (values.some((v) => !Number.isFinite(v))) return []
  const total = values.reduce((a, b) => a + b, 0)
  const out: string[] = []
  for (let i = 0; i < parts.length; i++) {
    // Parts are varied in their written 2-decimal form ("6.00", not "6"), or a
    // repair could only ever touch the digits before the point.
    for (const variant of digitConfusionVariants([values[i].toFixed(2)])) {
      const average = round2((total - values[i] + Number(variant)) / denominator)
      if (average <= 0) continue
      const canonical = String(average)
      if (!out.includes(canonical)) out.push(canonical)
    }
  }
  return out
}

/**
 * Digits this model swaps for each other on handwriting: a written 0 with an
 * open top reads as 6, a flat-topped 5 as 3 or 6, a barred 7 as 1, and so on.
 */
const DIGIT_CONFUSIONS: Record<string, string[]> = {
  '0': ['6', '8', '9', '4'],
  '1': ['7', '4'],
  '2': ['7'],
  '3': ['5', '8'],
  '4': ['9', '1', '0'],
  '5': ['3', '6', '8'],
  '6': ['0', '5', '8'],
  '7': ['1', '2'],
  '8': ['0', '3', '5', '6'],
  '9': ['0', '4']
}

/**
 * Variants of `values` with one digit shape corrected — either at a single
 * position, or at every position holding that digit at once. The all-positions
 * form matters because one clerk's habitual 0 is read as 6 everywhere in the
 * same figure ("15.66" for 15.00), which no single-position swap recovers.
 */
export function digitConfusionVariants(values: string[]): string[] {
  const out: string[] = []
  const push = (v: string): void => {
    const n = Number(v)
    if (!Number.isFinite(n) || n <= 0) return
    const canonical = String(n)
    if (values.some((v) => Number(v) === n) || out.some((v) => Number(v) === n)) return
    out.push(canonical)
  }
  for (const value of values) {
    for (let i = 0; i < value.length; i++) {
      const alternatives = DIGIT_CONFUSIONS[value[i]]
      if (alternatives === undefined) continue
      for (const alt of alternatives) {
        push(value.slice(0, i) + alt + value.slice(i + 1))
        push(value.split(value[i]).join(alt))
      }
    }
  }
  return out
}

function averageOf(numerator: string, count: number): string | undefined {
  const parts = splitAverageNumerator(numerator, count)
  if (parts === undefined) return undefined
  const total = parts.reduce((a, b) => a + Number(b), 0)
  return String(round2(total / count))
}

/** Every figure on this form is written to 2 decimals; display them that way. */
export function formatMbFigure(value: string): string {
  if (value === '') return ''
  const n = Number(value)
  return Number.isFinite(n) ? n.toFixed(2) : value
}

/**
 * The "No." column as written — "1x1", or a plain count — with letter-shaped
 * digits repaired ("IXI", "lX)", "(x1" all being 1x1). Blank when the cell
 * holds no count at all.
 */
export function readMbCount(raw: string): string {
  // Same guard as readMbFigure: with no real digit in the cell, this is OCR
  // noise from a blank column ("pwell"), and mapping its letter shapes would
  // turn it into a count of 11.
  if (!/[0-9]/.test(raw)) return ''
  const cleaned = raw
    .trim()
    .split('')
    .map((ch) => LETTER_DIGITS[ch] ?? ch)
    .join('')
  const pair = cleaned.match(/([0-9]+)\s*[xX×*]\s*([0-9]+)/)
  if (pair) return `${Number(pair[1])}x${Number(pair[2])}`
  // A multiplier whose "x" was read as a letter or lost leaves two 1-digit
  // runs ("11", "1 1"); a single run is a plain count.
  const single = cleaned.match(/[0-9]+/)
  return single ? String(Number(single[0])) : ''
}

export function parseMbCount(raw: string): number {
  const cleaned = raw
    .trim()
    .split('')
    .map((ch) => LETTER_DIGITS[ch] ?? ch)
    .join('')
  const factors = cleaned.match(/[0-9]+(?:\.[0-9]+)?/g)
  if (!factors || factors.length === 0) return 1
  const product = factors.map(Number).reduce((a, b) => a * b, 1)
  return Number.isFinite(product) && product > 0 ? product : 1
}

export type MbRowCheck = 'OK' | 'Corrected' | 'Recheck' | ''

export interface MbRowReconciliation {
  /** Corrected figures, best reading first — empty when that cell had none. */
  l: string
  b: string
  d: string
  contents: string
  /** No x L x B x D from the (corrected) dimensions, blank when not computable. */
  computed: string
  check: MbRowCheck
}

/** Rounds to the 2 decimals every figure on this form is written to. */
function round2(n: number): number {
  return Math.round(n * 100) / 100
}

/** Books round the product themselves, so allow a cent of rounding slack. */
const CONTENTS_TOLERANCE = 0.02

/** Cost of accepting a digit-repair variant instead of a reading the model returned. */
const VARIANT_COST = 20

/** How far a solved figure may stray from Contents: the book's own 2-decimal rounding, no more. */
const ROUNDING_TOLERANCE = 0.0051

/** One possible reading of a cell, with what it costs to believe it. */
interface Reading {
  value: string
  cost: number
  repaired: boolean
}

/**
 * Reconcile one row's figures against its own arithmetic.
 *
 * Picks the combination of candidate readings whose product matches the
 * Contents column, preferring each cell's best (first) candidate. When no
 * combination agrees, the figures are left as read and the row is marked
 * "Recheck" with the computed value alongside — that is either an OCR miss on
 * a badly written figure or a real error in the book, and both need a human.
 */
/**
 * Last resort when no combination of readings multiplies out to the Contents
 * figure: treat one dimension as unknown and solve it from the other two and
 * Contents. A clerk's row is arithmetic the book itself already checked, so a
 * value recovered this way is usually the one the page actually shows — but
 * only when it lands on a clean 2-decimal figure that reproduces Contents,
 * which a wrong guess almost never does.
 */
function solveOneDimension(
  count: number,
  options: string[][],
  contentsCandidates: string[],
  alternates: string[][],
  ranked: Reading[][],
  read: boolean[]
): { picks: string[]; contents: string } | undefined {
  let best: { picks: string[]; contents: string; cost: number } | undefined
  for (let slot = 0; slot < 3; slot++) {
    const others = [0, 1, 2].filter((i) => i !== slot)
    // The two cells the third is solved from stay close to what was read —
    // their strongest few readings, plus at most one digit repair between them
    // — or the search finds an arithmetic coincidence instead of the real row.
    // A dimension is only ever solved from two cells that were genuinely
    // read. Solving from a blank cell (silently taken as 1) or from ink the
    // model mangled produces a figure that satisfies the arithmetic and means
    // nothing.
    if (!read[others[0]] || !read[others[1]]) continue
    const choices = (i: number): Reading[] =>
      ranked[i].filter((r, idx) => (r.repaired ? true : idx < 3)).slice(0, 12)
    for (const oa of choices(others[0])) {
      for (const ob of choices(others[1])) {
        if (oa.repaired && ob.repaired) continue
        const va = oa.value
        const vb = ob.value
        const known = count * (va === '' ? 1 : Number(va)) * (vb === '' ? 1 : Number(vb))
        if (!Number.isFinite(known) || known <= 0) continue
        // Only the best reading of Contents anchors a solve. Contents is the
        // figure everything else is derived from here, and its weaker
        // candidates (the same digits with the point moved, "3015" for 30.15)
        // will happily support a solved dimension of 667.33.
        {
          const c = 0
          const contents = Number(contentsCandidates[c])
          if (!Number.isFinite(contents) || contents <= 0) continue
          const solved = round2(contents / known)
          // Reject a solved value outside the range this form records — a
          // length, breadth or depth in metres — or one that does not
          // reproduce Contents once rounded to the book's 2 decimals.
          if (solved < 0.01 || solved > 100) continue
          // A solved value is held to a tighter tolerance than a read one: it
          // must reproduce Contents to within the book's own 2-decimal
          // rounding. A looser fit means the figures it was solved from are
          // themselves wrong, and the "solution" is arithmetic noise.
          if (Math.abs(known * solved - contents) > ROUNDING_TOLERANCE) continue
          const picks = ['', '', '']
          picks[others[0]] = va
          picks[others[1]] = vb
          picks[slot] = String(solved)
          // Prefer solutions built on the strongest readings, and prefer
          // agreeing with the solved cell's own candidates when it has any.
          const agrees =
            options[slot].some((v) => Number(v) === solved) || alternates[slot].some((v) => Number(v) === solved)
          // A cell that was genuinely read must be repaired into one of its
          // own readings or digit-repair variants. Only a cell with nothing
          // usable in it may take the value the arithmetic alone produces.
          if (!agrees && read[slot]) continue
          // Residual dominates every other term: the figures the clerk
          // actually wrote multiply out exactly, so among repairs that all
          // satisfy the rounding tolerance, the exact one is the row. Reading
          // ranks only break ties between equally exact solutions.
          const residual = Math.abs(known * solved - contents)
          const cost = residual * 100000 + oa.cost + ob.cost + c + (agrees ? 0 : 10)
          if (best === undefined || cost < best.cost) best = { picks, contents: contentsCandidates[c], cost }
        }
      }
    }
  }
  return best === undefined ? undefined : { picks: best.picks, contents: best.contents }
}

export function reconcileMbRow(input: {
  /** How many times the measurement repeats — the "No." column's multiplier. */
  count: number
  /** Candidate readings per column, best first (see mbDimensionCandidates / mbFigureCandidates). */
  l: string[]
  b: string[]
  d: string[]
  contents: string[]
  /** Digit-repair candidates per dimension (see mbDimensionVariants), used only when nothing reconciles. */
  alternates?: { l: string[]; b: string[]; d: string[] }
  /** Per dimension: false when the cell's ink was read as something that cannot be a figure off this form. */
  trusted?: { l: boolean; b: boolean; d: boolean }
}): MbRowReconciliation {
  const count = input.count
  const dims = [input.l, input.b, input.d]
  const contentsCandidates = input.contents

  const best = (i: number): string => dims[i][0] ?? ''
  const fallback: MbRowReconciliation = {
    l: best(0),
    b: best(1),
    d: best(2),
    contents: contentsCandidates[0] ?? '',
    computed: '',
    check: ''
  }

  const hasDims = dims.some((c) => c.length > 0)
  if (!hasDims) return fallback

  const productOf = (picks: string[]): number =>
    picks.reduce((acc, v) => acc * (v === '' ? 1 : Number(v)), count)

  const computedFallback = round2(productOf([best(0), best(1), best(2)]))
  fallback.computed = String(computedFallback)

  if (contentsCandidates.length === 0) {
    // Nothing to check against: the computed value is all the row offers.
    return fallback
  }

  let bestMatch: { picks: string[]; contents: string; cost: number } | undefined
  const options = dims.map((c) => (c.length > 0 ? c : ['']))
  const alternates = [input.alternates?.l ?? [], input.alternates?.b ?? [], input.alternates?.d ?? []]
  // Every reading of a cell, cheapest first: what the model actually returned,
  // then digit-repair variants at a cost no stack of direct readings can reach.
  // A variant therefore only ever wins when no combination of real readings
  // multiplies out to Contents.
  const ranked: Reading[][] = options.map((values, i) => [
    ...values.map((value, idx) => ({ value, cost: idx, repaired: false })),
    ...alternates[i].map((value, idx) => ({ value, cost: VARIANT_COST + idx, repaired: true }))
  ])
  for (const a of ranked[0]) {
    for (const b of ranked[1]) {
      for (const d of ranked[2]) {
        // At most one repaired cell per row. With two or three cells free to
        // change digits, some combination almost always multiplies out to
        // Contents by coincidence, and the row is then silently rewritten into
        // figures the page does not show.
        if ([a, b, d].filter((o) => o.repaired).length > 1) continue
        const picks = [a.value, b.value, d.value]
        const product = productOf(picks)
        for (let c = 0; c < contentsCandidates.length; c++) {
          if (Math.abs(product - Number(contentsCandidates[c])) > CONTENTS_TOLERANCE) continue
          const cost = a.cost + b.cost + d.cost + c
          if (bestMatch === undefined || cost < bestMatch.cost) {
            bestMatch = { picks, contents: contentsCandidates[c], cost }
          }
        }
      }
    }
  }

  if (bestMatch === undefined) {
    // A cell counts as read only when the model returned something that could
    // be a figure off this form AND the caller did not flag the ink itself as
    // unreadable.
    const read = options.map((values, i) => {
      const bestRead = values[0] ?? ''
      const plausible = bestRead !== '' && Number(bestRead) <= 1000 && bestRead.replace(/[^0-9]/g, '').length <= 4
      const trusted = [input.trusted?.l, input.trusted?.b, input.trusted?.d][i] ?? true
      return plausible && trusted
    })
    const solved = solveOneDimension(count, options, contentsCandidates, alternates, ranked, read)
    if (solved !== undefined) {
      return {
        l: solved.picks[0],
        b: solved.picks[1],
        d: solved.picks[2],
        contents: solved.contents,
        computed: String(round2(productOf(solved.picks))),
        check: 'Corrected'
      }
    }
    return { ...fallback, check: 'Recheck' }
  }
  return {
    l: bestMatch.picks[0],
    b: bestMatch.picks[1],
    d: bestMatch.picks[2],
    contents: bestMatch.contents,
    computed: String(round2(productOf(bestMatch.picks))),
    check: bestMatch.cost === 0 ? 'OK' : 'Corrected'
  }
}

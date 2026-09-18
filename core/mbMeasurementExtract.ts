/**
 * MB (Measurement Book) "L.F. No. 83" sheet: a fixed preprinted form with
 * columns Date | Description of Work | No. | L | B | D | Contents. Every
 * page of every MB uses the identical stationery, so column positions and
 * row pitch are stable constants rather than something re-derived per page
 * (see electron/mbMeasurementOcr.ts for the pixel-grid detection that uses
 * these as its calibration reference).
 */

import {
  formatMbFigure,
  mbDimensionCandidates,
  mbAverageVariants,
  mbDimensionVariants,
  mbFigureCandidates,
  mbMappedDigitCount,
  parseMbCount,
  readMbCount,
  readMbFigure,
  reconcileMbRow,
  splitAverageNumerator
} from './mbNumeric'
import type { MbRowCheck } from './mbNumeric'

export const MB_MEASUREMENT_COLUMNS = ['Date', 'Description', 'No', 'L', 'B', 'D', 'Contents'] as const
export type MbColumn = (typeof MB_MEASUREMENT_COLUMNS)[number]

export interface MbMeasurementRow {
  date: string
  description: string
  no: string
  l1: string
  l2: string
  l3: string
  b1: string
  b2: string
  b3: string
  d1: string
  d2: string
  d3: string
  contents: string
  /** No. x L x B x D from this row's own figures — the book's arithmetic, recomputed. */
  computed: string
  /** '' when the row can't be checked, 'OK'/'Corrected' when the arithmetic agrees, 'Recheck' when it doesn't. */
  check: MbRowCheck
}

export const MB_MEASUREMENT_HEADERS = [
  'Date',
  'Description of Work',
  'No.',
  'L1',
  'L2',
  'L3',
  'B1',
  'B2',
  'B3',
  'D1',
  'D2',
  'D3',
  'Contents',
  'Computed',
  'Check'
] as const

/**
 * Estimate the fixed row pitch (px between consecutive ruled lines) from a
 * set of detected horizontal line y-positions. Only gaps within a plausible
 * row-height range are averaged, so unrelated title-block rules (whose
 * spacing doesn't match the grid) don't skew the estimate.
 */
export function estimateRowPitch(lineYs: number[], opts: { minGap?: number; maxGap?: number } = {}): number | undefined {
  const minGap = opts.minGap ?? 30
  const maxGap = opts.maxGap ?? 70
  const gaps: number[] = []
  for (let i = 1; i < lineYs.length; i++) {
    const g = lineYs[i] - lineYs[i - 1]
    if (g >= minGap && g <= maxGap) gaps.push(g)
  }
  if (gaps.length === 0) return undefined
  return gaps.reduce((a, b) => a + b, 0) / gaps.length
}

/**
 * Row pitch from the dominant gap, plus how many gaps agree with it and how
 * far the worst of those strays. A page photographed at a different
 * distance/crop than the calibration reference has its own valid pitch that
 * doesn't match REF_ROW_PITCH's scaled value — so detected spacing should be
 * trusted when it's internally consistent (many gaps that agree with each
 * other), not just when it happens to match the reference.
 *
 * The gaps are clustered rather than averaged: the header block above the
 * measurement grid (the two-line column captions) rules at its own, smaller
 * spacing, and those gaps can sit inside the plausible range too. Averaging
 * them together with the real row gaps drags the pitch down by enough that
 * generated row bounds slice through the handwriting a few rows down.
 */
export function estimateRowPitchConfidence(
  lineYs: number[],
  opts: { minGap?: number; maxGap?: number } = {}
): { pitch: number; sampleCount: number; maxDeviation: number } | undefined {
  const minGap = opts.minGap ?? 30
  const maxGap = opts.maxGap ?? 70
  const all: number[] = []
  for (let i = 1; i < lineYs.length; i++) {
    const g = lineYs[i] - lineYs[i - 1]
    if (g >= minGap && g <= maxGap) all.push(g)
  }
  const gaps = dominantGapCluster(all)
  if (gaps.length === 0) return undefined
  const pitch = gaps.reduce((a, b) => a + b, 0) / gaps.length
  const maxDeviation = Math.max(...gaps.map((g) => Math.abs(g - pitch)))
  return { pitch, sampleCount: gaps.length, maxDeviation }
}

/**
 * The largest group of gaps that agree with each other (within 10%, which
 * comfortably covers photo perspective across one page while still separating
 * header spacing from row spacing). Ties go to the wider spacing, since the
 * measurement rows are the widest repeated rule on the form.
 */
function dominantGapCluster(gaps: number[]): number[] {
  let best: number[] = []
  for (const centre of gaps) {
    const cluster = gaps.filter((g) => Math.abs(g - centre) <= centre * 0.1)
    const better =
      cluster.length > best.length ||
      (cluster.length === best.length && best.length > 0 && mean(cluster) > mean(best))
    if (better) best = cluster
  }
  return best
}

function mean(values: number[]): number {
  return values.reduce((a, b) => a + b, 0) / values.length
}

/**
 * First line-Y where the fixed-pitch measurement grid begins: the first line
 * whose gap to the FOLLOWING line is within tolerance of the estimated pitch.
 * Skips the form's title-block rules, whose spacing doesn't match the grid.
 */
export function findGridStart(lineYs: number[], pitch: number, tolerance = 6): number | undefined {
  for (let i = 0; i < lineYs.length - 1; i++) {
    if (Math.abs(lineYs[i + 1] - lineYs[i] - pitch) <= tolerance) return lineYs[i]
  }
  return undefined
}

/** Row boundaries generated arithmetically from a start Y and fixed pitch, up to maxRows bands or bottomLimit, whichever comes first. */
export function generateRowBounds(start: number, pitch: number, bottomLimit: number, maxRows = 20): number[] {
  const bounds: number[] = [start]
  let y = start
  for (let i = 0; i < maxRows && y + pitch <= bottomLimit + pitch * 0.5; i++) {
    y += pitch
    bounds.push(Math.round(y))
  }
  return bounds
}

/**
 * Row boundaries generated the same way as generateRowBounds (start Y +
 * fixed pitch, up to bottomLimit), except each step snaps to the nearest
 * actually-detected line when one lands close to the expected position.
 * A single averaged pitch is never exactly right for a real photo (a couple
 * of stray gaps pull the average off by a few px), so pure arithmetic
 * stepping drifts further from the true grid with every row. Snapping to a
 * real detected line each time it's available resets that drift; a row
 * whose line got missed (handwriting occluding it) just falls back to the
 * arithmetic estimate for that one step, same as generateRowBounds.
 */
export function generateRowBoundsSnapped(
  detectedRows: number[],
  start: number,
  pitch: number,
  bottomLimit: number,
  maxRows = 20,
  snapTolerance = pitch * 0.3
): number[] {
  const sorted = [...detectedRows].sort((a, b) => a - b)
  const bounds: number[] = [start]
  let y = start
  for (let i = 0; i < maxRows && y + pitch <= bottomLimit + pitch * 0.5; i++) {
    const expected = y + pitch
    let nearest: number | undefined
    let nearestDist = Infinity
    for (const r of sorted) {
      const d = Math.abs(r - expected)
      if (d < nearestDist) {
        nearestDist = d
        nearest = r
      }
    }
    y = nearest !== undefined && nearestDist <= snapTolerance ? nearest : expected
    bounds.push(Math.round(y))
  }
  return bounds
}

/**
 * Pick the 8 true column-divider x-positions out of a noisier set of
 * detected peaks. Raw peak detection on a real photo often includes extra
 * candidates from handwritten fraction-bar strokes that happen to line up
 * across rows (e.g. a recurring "6.70+6.70" numerator), alongside the real
 * ruled lines. No/L/B/D are printed as 4 equal-width columns on this form
 * (true of the reference and confirmed on real photos), so the correct
 * subset of 8 is the one whose middle 4 gaps (No, L, B, D widths) are most
 * consistent with each other — a wrong subset produces wildly uneven gaps.
 * The 2 outer edges (leftmost/rightmost peak) are trusted as fixed anchors:
 * the table border is normally the strongest, most complete line on the
 * page. Returns undefined if fewer than 8 peaks are available, or if no
 * subset achieves a tight enough spread to be trusted.
 */
export function selectColumnBoundsByEqualWidth(
  peaks: number[],
  opts: { maxSpreadRatio?: number } = {}
): number[] | undefined {
  const maxSpreadRatio = opts.maxSpreadRatio ?? 0.2
  const sorted = [...peaks].sort((a, b) => a - b)
  if (sorted.length < 8) return undefined
  if (sorted.length === 8) {
    const spread = middleGapSpreadRatio(sorted)
    return spread !== undefined && spread <= maxSpreadRatio ? sorted : undefined
  }

  const first = sorted[0]
  const last = sorted[sorted.length - 1]
  const interior = sorted.slice(1, -1)

  let best: number[] | undefined
  let bestSpread = Infinity
  for (const combo of combinations(interior, 6)) {
    const candidate = [first, ...combo, last]
    const spread = middleGapSpreadRatio(candidate)
    if (spread !== undefined && spread < bestSpread) {
      bestSpread = spread
      best = candidate
    }
  }
  return best !== undefined && bestSpread <= maxSpreadRatio ? best : undefined
}

function middleGapSpreadRatio(bounds: number[]): number | undefined {
  const gaps = [bounds[3] - bounds[2], bounds[4] - bounds[3], bounds[5] - bounds[4], bounds[6] - bounds[5]]
  if (gaps.some((g) => g <= 0)) return undefined
  const mean = gaps.reduce((a, b) => a + b, 0) / gaps.length
  const range = Math.max(...gaps) - Math.min(...gaps)
  return range / mean
}

function* combinations<T>(items: T[], k: number): Generator<T[]> {
  if (k === 0) {
    yield []
    return
  }
  for (let i = 0; i <= items.length - k; i++) {
    for (const rest of combinations(items.slice(i + 1), k - 1)) {
      yield [items[i], ...rest]
    }
  }
}

/**
 * Leading numeric-looking token in an OCR'd No./L/B/D/Contents cell: a plain
 * number, a "1x1"-style multiplier, or a dash placeholder. Cells in this
 * column group that DON'T start with one of these (e.g. "pwell", "CeoL") are
 * OCR noise hallucinated from a blank or near-blank crop — general OCR
 * models tend to produce plausible-looking text rather than nothing when fed
 * paper grain/shadow — and are dropped rather than shown as fabricated data.
 */
const LEADING_NUMERIC_TOKEN = /^\s*([0-9]+(?:\.[0-9]+)?(?:\s*[xX×]\s*[0-9]+(?:\.[0-9]+)?)?|[—–-])/

export function extractNumericToken(raw: string): string {
  const m = raw.match(LEADING_NUMERIC_TOKEN)
  return m ? m[1].replace(/\s+/g, '') : ''
}

/**
 * Split a raw L/B/D cell's text into up to 3 values. Multiple stacked
 * measurements in one cell (e.g. a wall measured in 3 segments) come through
 * as separate lines/segments from OCR; a single measurement lands in the
 * first slot only, leaving the other two blank.
 */
export function splitMultiValue(raw: string): [string, string, string] {
  const parts = raw
    .split(/\r?\n/)
    .map((s) => readMbFigure(s))
    .filter((s) => s.length > 0)
  return [parts[0] ?? '', parts[1] ?? '', parts[2] ?? '']
}

export function normalizeNumericCell(raw: string): string {
  const s = raw.trim()
  if (!s) return s
  // Decimal separator sometimes misread as a comma / colon / full-width comma.
  return s.replace(/[、;]/g, '.').replace(/:/g, '.')
}

/**
 * Any of L/B/D on this form can be filled as the average of two or three
 * readings, written as a small handwritten fraction: numerator "a+b" or
 * "a+b+c" over a denominator (e.g. "6.70+6.70" over "2", or "6.66+7.50+7.50"
 * over "3"). OCR renders the fraction either as two stacked lines (sum, then
 * denominator) or as one line with a literal "/" where the fraction bar was.
 * This recognizes both shapes and resolves them to the averaged value
 * instead of the fraction being torn apart into unrelated readings by
 * splitMultiValue.
 */
export function resolveAverageFraction(raw: string): string | undefined {
  const lines = raw
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0)

  let sumText: string | undefined
  let denominatorText: string | undefined
  if (lines.length === 2) {
    ;[sumText, denominatorText] = lines
  } else if (lines.length === 1) {
    const m = lines[0].match(/^(.+?)\/\s*(\d+(?:\.\d+)?)\s*$/)
    if (m) [sumText, denominatorText] = [m[1], m[2]]
  }
  if (sumText === undefined || denominatorText === undefined) return undefined

  const sum = sumText.replace(/\s+/g, '')
  if (!/^\d+(?:\.\d+)?(?:\+\d+(?:\.\d+)?)+$/.test(sum)) return undefined
  const denominator = Number(denominatorText)
  if (!Number.isFinite(denominator) || denominator <= 0) return undefined

  const total = sum.split('+').reduce((a, b) => a + Number(b), 0)
  return String(Math.round((total / denominator) * 100) / 100)
}

/**
 * An averaged cell resolved to both its value and the individual figures it
 * averages, in either the clean or the mangled shape. The parts are what a
 * digit repair has to work on (see mbAverageVariants).
 */
function resolveAverageParts(raw: string): { value: string; parts: string[]; denominator: number } | undefined {
  const lines = raw
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
  const clean = resolveAverageFraction(raw)
  if (clean !== undefined) {
    const single = lines.length === 1 ? lines[0].match(/^(.+?)\/\s*(\d+(?:\.\d+)?)\s*$/) : undefined
    const sumText = single ? single[1] : lines[0]
    const denominator = Number(single ? single[2] : lines[1])
    return { value: clean, parts: sumText.replace(/\s+/g, '').split('+'), denominator }
  }
  const mangled = resolveAverageCell(raw)
  if (mangled === undefined) return undefined
  const denominator = parseMbCount(lines[1] ?? '')
  const parts = splitAverageNumerator(lines[0] ?? '', denominator)
  if (parts === undefined) return undefined
  return { value: mangled, parts, denominator }
}

/**
 * cells columns must follow MB_MEASUREMENT_COLUMNS order: [date, description,
 * no, l, b, d, contents].
 *
 * Read in two passes. The first reconciles each row on its own. The second
 * re-reads the rows that failed, this time knowing the depth the page's
 * verified rows agree on — an item is measured at one depth down the whole
 * block, so that value is evidence from the page itself rather than a guess,
 * and it commonly rescues a row where the depth cell was misread ("630" for
 * 0.30) badly enough to break the arithmetic.
 */
export function buildMbMeasurementRows(grid: string[][]): MbMeasurementRow[] {
  const firstPass = readRows(grid, undefined)
  const depthPrior = verifiedDepth(firstPass)
  return depthPrior === undefined ? firstPass : readRows(grid, depthPrior)
}

/**
 * The depth that rows whose arithmetic checked out agree on, if they do. Only
 * rows verified against their own Contents figure count as evidence — an
 * unverified row's depth is exactly the reading in doubt.
 */
function verifiedDepth(rows: MbMeasurementRow[]): string | undefined {
  const verified = rows.filter((r) => r.check === 'OK' && r.d1 !== '' && r.d2 === '').map((r) => r.d1)
  if (verified.length < 2) return undefined
  const counts = new Map<string, number>()
  for (const v of verified) counts.set(v, (counts.get(v) ?? 0) + 1)
  const [value, count] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]
  return count >= verified.length / 2 ? value : undefined
}

/**
 * Whether a dimension cell's OCR text can be a figure off this form at all.
 * Every figure here is written to 2 decimals, so 3 digits: fewer means most of
 * the ink was missed ("3" where 7.20 is written), and more means several
 * figures ran together ("6046-40" over "2846") — unless the cell is a written
 * average, which legitimately carries one group of 3 digits per part.
 */
function isReadableCell(raw: string): boolean {
  const digits = mbMappedDigitCount(raw)
  if (digits === 0) return false
  if (digits >= 3 && digits <= 4) return true
  if (resolveAverageParts(raw) !== undefined) return true
  // An average whose fraction bar was lost still reads as a figure per part,
  // but only when the "+" signs survived — without them, a long digit run is
  // just as likely to be two rows' ink read as one cell.
  const numerator = raw.split(/\r?\n/)[0] ?? ''
  if (!numerator.includes('+')) return false
  return [2, 3].some((n) => splitAverageNumerator(numerator, n) !== undefined)
}

/** One L/B/D cell resolved into what the row's arithmetic may choose between. */
interface Dimension {
  candidates: string[]
  stacked: [string, string]
  alternates: string[]
  /** Set only for a stacked cell: every line read as a candidate for the single dimension. */
  lineCandidates?: string[]
  lineAlternates?: string[]
}

function readRows(grid: string[][], depthPrior: string | undefined): MbMeasurementRow[] {
  return grid.map((cells) => {
    /**
     * A dimension cell yields either one value with candidate readings for
     * the row's arithmetic to choose between, or — when the clerk stacked
     * several sub-measurements in the cell — the stacked values as read.
     */
    const resolveDimension = (raw: string): Dimension => {
      const averaged = resolveAverageParts(raw)
      if (averaged !== undefined) {
        return {
          candidates: [averaged.value],
          stacked: ['', ''],
          alternates: mbAverageVariants(averaged.parts, averaged.denominator)
        }
      }
      const [first, second, third] = splitMultiValue(raw)
      if (second !== '') {
        // Every line offered as a reading of this one dimension as well: the
        // extra line is as often a neighbouring row's figure leaning into the
        // cell as a genuine second sub-measurement, and only the row's
        // arithmetic can tell those apart (see the collapse below).
        const lines = [first, second, third].filter((v) => v !== '')
        return {
          candidates: first === '' ? [] : [first],
          stacked: [second, third],
          alternates: [],
          lineCandidates: lines.flatMap((v) => mbDimensionCandidates(v)),
          lineAlternates: lines.flatMap((v) => mbDimensionVariants(v))
        }
      }
      return { candidates: mbDimensionCandidates(raw), stacked: ['', ''], alternates: mbDimensionVariants(raw) }
    }
    const l = resolveDimension(cells[3] ?? '')
    const b = resolveDimension(cells[4] ?? '')
    const d = resolveDimension(cells[5] ?? '')
    if (depthPrior !== undefined) {
      if (d.candidates.length > 0 && !d.candidates.includes(depthPrior)) d.candidates = [...d.candidates, depthPrior]
      if (d.lineCandidates !== undefined && !d.lineCandidates.includes(depthPrior)) {
        d.lineCandidates = [...d.lineCandidates, depthPrior]
      }
    }
    const no = readMbCount(cells[2] ?? '')
    const contentsCandidates = mbFigureCandidates(cells[6] ?? '')
    // A row measured as a single L/B/D triple is checked against its own
    // arithmetic. A row with stacked sub-measurements is checked too, reading
    // each stacked line as a candidate for the one dimension: when that makes
    // the row multiply out to Contents, the "sub-measurements" were really one
    // figure plus a neighbouring row's ink, and the cell collapses to the
    // figure the arithmetic proves. When it doesn't, the stacked values stand
    // as read — the clerk combines genuine sub-measurements in ways the sheet
    // doesn't record, and this tool must not invent that working.
    const stacked = [l, b, d].some((dim) => dim.stacked[0] !== '')
    const pick = (dim: Dimension): string[] => (stacked ? dim.lineCandidates ?? dim.candidates : dim.candidates)
    const pickAlt = (dim: Dimension): string[] => (stacked ? dim.lineAlternates ?? dim.alternates : dim.alternates)
    const reconciled = reconcileMbRow({
      count: parseMbCount(no),
      l: pick(l),
      b: pick(b),
      d: pick(d),
      contents: contentsCandidates,
      alternates: { l: pickAlt(l), b: pickAlt(b), d: pickAlt(d) },
      trusted: { l: isReadableCell(cells[3] ?? ''), b: isReadableCell(cells[4] ?? ''), d: isReadableCell(cells[5] ?? '') }
    })
    const collapsed = stacked && (reconciled.check === 'OK' || reconciled.check === 'Corrected')
    const row = collapsed || !stacked ? reconciled : undefined
    const sub = (dim: Dimension, i: number): string => (collapsed ? '' : formatMbFigure(dim.stacked[i]))
    const figure = (value: string | undefined, fallback: string[]): string =>
      formatMbFigure(value ?? fallback[0] ?? '')
    return {
      date: (cells[0] ?? '').trim(),
      description: (cells[1] ?? '').trim(),
      no,
      l1: figure(row?.l, l.candidates),
      l2: sub(l, 0),
      l3: sub(l, 1),
      b1: figure(row?.b, b.candidates),
      b2: sub(b, 0),
      b3: sub(b, 1),
      d1: figure(row?.d, d.candidates),
      d2: sub(d, 0),
      d3: sub(d, 1),
      contents: figure(row?.contents, contentsCandidates),
      computed: formatMbFigure(row?.computed ?? ''),
      check: row?.check ?? ''
    }
  })
}

/**
 * An averaged L/B/D cell whose "+" signs or fraction bar were mangled by OCR
 * ("6.3046-30" over "2" for 6.30+6.30 over 2). resolveAverageFraction handles
 * the clean shapes; this recovers the mangled ones structurally, using the
 * denominator to say how many 2-decimal figures the numerator must hold.
 */
function resolveAverageCell(raw: string): string | undefined {
  const lines = raw
    .split(/\r?\n/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0)
  if (lines.length !== 2) return undefined
  const count = parseMbCount(lines[1])
  if (!Number.isInteger(count) || count < 2 || count > 4) return undefined
  const parts = splitAverageNumerator(lines[0], count)
  if (parts === undefined) return undefined
  const total = parts.reduce((a, b) => a + Number(b), 0)
  return String(Math.round((total / count) * 100) / 100)
}

export function mbMeasurementRowsToGrid(rows: MbMeasurementRow[]): string[][] {
  return [
    [...MB_MEASUREMENT_HEADERS],
    ...rows.map((r) => [
      r.date,
      r.description,
      r.no,
      r.l1,
      r.l2,
      r.l3,
      r.b1,
      r.b2,
      r.b3,
      r.d1,
      r.d2,
      r.d3,
      r.contents,
      r.computed,
      r.check
    ])
  ]
}

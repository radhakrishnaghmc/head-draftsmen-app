/**
 * MB (Measurement Book) "L.F. No. 83" sheet: a fixed preprinted form with
 * columns Date | Description of Work | No. | L | B | D | Contents. Every
 * page of every MB uses the identical stationery, so column positions and
 * row pitch are stable constants rather than something re-derived per page
 * (see electron/mbMeasurementOcr.ts for the pixel-grid detection that uses
 * these as its calibration reference).
 */

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
  'Contents'
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
 * Same gap estimate as estimateRowPitch, plus how many gaps it's built from
 * and how far the worst of them strays from the average. A page photographed
 * at a different distance/crop than the calibration reference has its own
 * valid pitch that doesn't match REF_ROW_PITCH's scaled value — so detected
 * spacing should be trusted when it's internally consistent (many gaps that
 * agree with each other), not just when it happens to match the reference.
 */
export function estimateRowPitchConfidence(
  lineYs: number[],
  opts: { minGap?: number; maxGap?: number } = {}
): { pitch: number; sampleCount: number; maxDeviation: number } | undefined {
  const minGap = opts.minGap ?? 30
  const maxGap = opts.maxGap ?? 70
  const gaps: number[] = []
  for (let i = 1; i < lineYs.length; i++) {
    const g = lineYs[i] - lineYs[i - 1]
    if (g >= minGap && g <= maxGap) gaps.push(g)
  }
  if (gaps.length === 0) return undefined
  const pitch = gaps.reduce((a, b) => a + b, 0) / gaps.length
  const maxDeviation = Math.max(...gaps.map((g) => Math.abs(g - pitch)))
  return { pitch, sampleCount: gaps.length, maxDeviation }
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
    .split(/\r?\n|[,;]/)
    .map((s) => extractNumericToken(normalizeNumericCell(s)))
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

/** cells columns must follow MB_MEASUREMENT_COLUMNS order: [date, description, no, l, b, d, contents]. */
export function buildMbMeasurementRows(grid: string[][]): MbMeasurementRow[] {
  return grid.map((cells) => {
    const resolveDimension = (raw: string): [string, string, string] => {
      const averaged = resolveAverageFraction(raw)
      return averaged !== undefined ? [averaged, '', ''] : splitMultiValue(raw)
    }
    const [l1, l2, l3] = resolveDimension(cells[3] ?? '')
    const [b1, b2, b3] = resolveDimension(cells[4] ?? '')
    const [d1, d2, d3] = resolveDimension(cells[5] ?? '')
    return {
      date: (cells[0] ?? '').trim(),
      description: (cells[1] ?? '').trim(),
      no: extractNumericToken(normalizeNumericCell(cells[2] ?? '')),
      l1,
      l2,
      l3,
      b1,
      b2,
      b3,
      d1,
      d2,
      d3,
      contents: extractNumericToken(normalizeNumericCell(cells[6] ?? ''))
    }
  })
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
      r.contents
    ])
  ]
}

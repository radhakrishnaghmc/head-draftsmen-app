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
 * Split a raw L/B/D cell's text into up to 3 values. Multiple stacked
 * measurements in one cell (e.g. a wall measured in 3 segments) come through
 * as separate lines/segments from OCR; a single measurement lands in the
 * first slot only, leaving the other two blank.
 */
export function splitMultiValue(raw: string): [string, string, string] {
  const parts = raw
    .split(/\r?\n|[,;]/)
    .map((s) => normalizeNumericCell(s))
    .filter((s) => s.length > 0)
  return [parts[0] ?? '', parts[1] ?? '', parts[2] ?? '']
}

export function normalizeNumericCell(raw: string): string {
  const s = raw.trim()
  if (!s) return s
  // Decimal separator sometimes misread as a comma / colon / full-width comma.
  return s.replace(/[、;]/g, '.').replace(/:/g, '.')
}

/** cells columns must follow MB_MEASUREMENT_COLUMNS order: [date, description, no, l, b, d, contents]. */
export function buildMbMeasurementRows(grid: string[][]): MbMeasurementRow[] {
  return grid.map((cells) => {
    const [l1, l2, l3] = splitMultiValue(cells[3] ?? '')
    const [b1, b2, b3] = splitMultiValue(cells[4] ?? '')
    const [d1, d2, d3] = splitMultiValue(cells[5] ?? '')
    return {
      date: (cells[0] ?? '').trim(),
      description: (cells[1] ?? '').trim(),
      no: normalizeNumericCell(cells[2] ?? ''),
      l1,
      l2,
      l3,
      b1,
      b2,
      b3,
      d1,
      d2,
      d3,
      contents: normalizeNumericCell(cells[6] ?? '')
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

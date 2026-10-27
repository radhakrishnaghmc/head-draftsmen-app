/**
 * Layout-independent grid detection for the MB "L.F. No. 83" measurement
 * form. Unlike detectGrid() in mbGridDetect.ts (calibrated to one scan's
 * width/crop), this finds the table from the page's own long ruled lines, so
 * scans from different scanners/crops and slightly skewed pages work.
 *
 * Pure functions over a grayscale buffer so they can be tested without sharp.
 */
import { otsuThreshold } from './mbGridDetect'
import type { GrayImage } from './mbGridDetect'

export interface AutoGrid {
  colBounds: number[]
  rowBounds: number[]
  darkThreshold: number
  /** 0..1 — how closely the detected structure matches the form (8 column lines, equal No/L/B/D widths). */
  score: number
}

/**
 * Skew (degrees, + = page content rotated clockwise) that makes the ruled
 * lines most horizontal: the angle maximising the sum of squared per-row ink
 * counts. Works on a coarse subsample of dark pixels, so it is cheap.
 */
export function estimateSkewDegrees(img: GrayImage, darkThreshold: number, range = 4, step = 0.25): number {
  const { data, width, height } = img
  const sub = Math.max(1, Math.round(width / 500))
  const xs: number[] = []
  const ys: number[] = []
  for (let y = 0; y < height; y += sub) {
    for (let x = 0; x < width; x += sub) {
      if (data[y * width + x] < darkThreshold) {
        xs.push(x)
        ys.push(y)
      }
    }
  }
  if (xs.length === 0) return 0
  let best = 0
  let bestScore = -1
  const bins = new Float64Array(height + width + 2)
  for (let a = -range; a <= range + 1e-9; a += step) {
    const rad = (a * Math.PI) / 180
    const sin = Math.sin(rad)
    const cos = Math.cos(rad)
    bins.fill(0)
    for (let i = 0; i < xs.length; i++) {
      const yr = Math.round(ys[i] * cos - xs[i] * sin) + width
      bins[yr]++
    }
    let s = 0
    for (let i = 0; i < bins.length; i++) s += bins[i] * bins[i]
    if (s > bestScore) {
      bestScore = s
      best = a
    }
  }
  return best
}

/** Longest vertical run of dark pixels in each column, tolerating `gap` light pixels inside a run. */
export function longestVerticalRuns(img: GrayImage, thr: number, y0: number, y1: number, gap = 3): Int32Array {
  const { data, width } = img
  const out = new Int32Array(width)
  for (let x = 0; x < width; x++) {
    let best = 0
    let start = -1
    let last = -1
    for (let y = y0; y < y1; y++) {
      if (data[y * width + x] < thr) {
        if (start < 0 || y - last > gap + 1) start = y
        last = y
        const len = last - start + 1
        if (len > best) best = len
      }
    }
    out[x] = best
  }
  return out
}

/** Longest horizontal run of dark pixels in each row (within [x0,x1)), tolerating `gap` light pixels. */
export function longestHorizontalRuns(img: GrayImage, thr: number, x0: number, x1: number, gap = 3): Int32Array {
  const { data, width, height } = img
  const out = new Int32Array(height)
  for (let y = 0; y < height; y++) {
    let best = 0
    let start = -1
    let last = -1
    const base = y * width
    for (let x = x0; x < x1; x++) {
      if (data[base + x] < thr) {
        if (start < 0 || x - last > gap + 1) start = x
        last = x
        const len = last - start + 1
        if (len > best) best = len
      }
    }
    out[y] = best
  }
  return out
}

/** Collapse runs of adjacent indices whose value passes `min` to one centre each (weighted by value). */
export function peaksAbove(values: ArrayLike<number>, min: number, mergeWithin: number): number[] {
  const peaks: number[] = []
  let i = 0
  const n = values.length
  while (i < n) {
    if (values[i] >= min) {
      let j = i
      let last = i
      let wsum = 0
      let sum = 0
      while (j < n && (values[j] >= min || j - last <= mergeWithin)) {
        if (values[j] >= min) {
          wsum += values[j] * j
          sum += values[j]
          last = j
        }
        j++
      }
      peaks.push(Math.round(wsum / sum))
      i = j
    } else {
      i++
    }
  }
  return peaks
}

export interface FormFit {
  /** x of the No | L | B | D | Contents-left lines and the Contents right edge (6 numbers). */
  bounds: number[]
  u: number
  matched: number
}

/**
 * Fit the form's equal-width No/L/B/D columns to whichever vertical ruled
 * lines were actually found (any 3+ of the 6 numeric-column lines fix the
 * rest, so faint or broken dividers don't matter). Grid points are
 * x0 + k*u for k = 0..4 (No-left, L-left, B-left, D-left, Contents-left) and
 * Contents-right at x0 + (4 + c)*u with c ~ 1.0..1.5.
 */
export function fitNumericColumns(peaks: number[], imgWidth: number): FormFit | undefined {
  if (peaks.length < 2) return undefined
  let best: (FormFit & { score: number }) | undefined
  const uMin = imgWidth * 0.035
  const uMax = imgWidth * 0.13
  for (let a = 0; a < peaks.length; a++) {
    for (let b = a + 1; b < peaks.length; b++) {
      for (let k = 1; k <= 4; k++) {
        const u = (peaks[b] - peaks[a]) / k
        if (u < uMin || u > uMax) continue
        for (let ka = 0; ka <= 4; ka++) {
          const x0 = peaks[a] - ka * u
          if (x0 < 0 || x0 + 5 * u > imgWidth * 1.02) continue
          const tol = 0.14 * u
          let matched = 0
          for (let g = 0; g <= 4; g++) {
            if (peaks.some((p) => Math.abs(p - (x0 + g * u)) <= tol)) matched++
          }
          const rightPeak = peaks.find((p) => p >= x0 + 4.9 * u - tol && p <= x0 + 5.6 * u + tol)
          const hasLeft = peaks.some((p) => Math.abs(p - x0) <= tol)
          const hasDescEdge = peaks.some((p) => p < x0 - 1.8 * u)
          const score = matched + (rightPeak !== undefined ? 1 : 0) + (hasLeft ? 0.7 : 0) + (hasDescEdge ? 0.3 : 0)
          if (matched >= 2 && (!best || score > best.score)) {
            const contentsRight = rightPeak ?? x0 + 5.3 * u
            best = { bounds: [x0, x0 + u, x0 + 2 * u, x0 + 3 * u, x0 + 4 * u, contentsRight], u, matched, score }
          }
        }
      }
    }
  }
  return best
}

/** Row rules across the numeric columns, regularised to the dominant pitch (noise dropped, gaps repaired). */
export function rowBoundsFor(img: GrayImage, x0: number, x1: number): number[] {
  const runs = longestHorizontalRuns(img, 128, Math.round(x0), Math.round(x1), 6)
  const lines = peaksAbove(runs, (x1 - x0) * 0.35, 4)
  if (lines.length < 3) return lines
  const gaps: number[] = []
  for (let i = 1; i < lines.length; i++) gaps.push(lines[i] - lines[i - 1])
  const sortedGaps = [...gaps].sort((a, b) => a - b)
  // Dominant pitch: the smallest gap that at least 3 gaps agree with (within 12%) — underlines make small spurious gaps.
  let pitch = sortedGaps[Math.floor(sortedGaps.length / 2)]
  for (const g of sortedGaps) {
    if (g < 15) continue
    if (sortedGaps.filter((h) => Math.abs(h - g) <= g * 0.12).length >= 3) {
      pitch = g
      break
    }
  }
  // Keep lines that sit one pitch (+-25%) from a kept neighbour; drop the weaker of two lines closer than 0.6 pitch.
  const strength = (y: number) => runs[y] ?? 0
  const kept: number[] = []
  for (const y of lines) {
    const prev = kept[kept.length - 1]
    if (prev !== undefined && y - prev < pitch * 0.6) {
      if (strength(y) > strength(prev)) kept[kept.length - 1] = y
      continue
    }
    kept.push(y)
  }
  const ok = kept.filter((y, i) => {
    const near = (d: number) =>
      Math.abs(d - pitch) <= pitch * 0.25 ||
      (d > pitch * 1.6 && Math.abs(d / Math.round(d / pitch) - pitch) <= pitch * 0.2)
    return (i > 0 && near(y - kept[i - 1])) || (i < kept.length - 1 && near(kept[i + 1] - y))
  })
  const out: number[] = []
  for (const y of ok) {
    const prev = out[out.length - 1]
    if (prev !== undefined) {
      const gap = y - prev
      const n = Math.round(gap / pitch)
      if (n >= 2 && n <= 4) for (let j = 1; j < n; j++) out.push(Math.round(prev + (gap * j) / n))
    }
    out.push(y)
  }
  return out
}

/**
 * @param mask ink mask (0 = ink, 255 = paper) — see inkMask in electron/mbMeasurementOcr.ts
 */
export function detectGridAuto(mask: GrayImage): AutoGrid | undefined {
  const { width, height } = mask
  const vruns = longestVerticalRuns(mask, 128, 0, height, 5)
  const peaks = peaksAbove(vruns, height * 0.18, 4)
  const fit = fitNumericColumns(peaks, width)
  if (!fit || fit.matched < 3) return undefined
  // Date and Description widths vary by book, so only their outer edges are approximate.
  const descRight = fit.bounds[0]
  const left = peaks.filter((p) => p < descRight - 1.8 * fit.u)
  const descLeft = left.length ? left[left.length - 1] : Math.max(0, descRight - 3.2 * fit.u)
  const dateLeft = left.length > 1 ? left[left.length - 2] : Math.max(0, descLeft - 1.2 * fit.u)
  const colBounds = [dateLeft, descLeft, ...fit.bounds]
  const rows = rowBoundsFor(mask, fit.bounds[0], fit.bounds[4])
  if (rows.length < 4) return undefined
  return { colBounds, rowBounds: rows, darkThreshold: 128, score: fit.matched / 5 }
}

/**
 * Pixel-level grid detection for the MB "L.F. No. 83" measurement form:
 * finds the 8 column-divider x positions and the ruled row lines in a
 * grayscale page image. Kept here (not in electron/) so it can be exercised
 * against real sample photos from tests without an Electron/sharp runtime —
 * callers decode the image and hand over raw grayscale bytes.
 */
import {
  estimateRowPitchConfidence,
  findGridStart,
  generateRowBounds,
  generateRowBoundsSnapped,
  selectColumnBoundsByEqualWidth
} from './mbMeasurementExtract'

// Calibrated once from a clean sample page of the form. This is a FIXED
// preprinted form — the same stationery on every page of every MB — so these
// are used as the primary source of truth for column bounds and row pitch,
// not just a fallback: per-page pixel detection is noisy (handwriting
// overlapping ruled lines, perspective/lighting differences between photos of
// different pages) and is only trusted as a refinement when it lands close to
// this reference.
export const REF_WIDTH = 1275
export const REF_COLUMN_BOUNDS = [78, 231, 697, 796, 895, 994, 1095, 1234] // 8 boundaries -> 7 columns
export const REF_ROW_TOP = 258
export const REF_ROW_PITCH = 48.71

export interface GrayImage {
  data: Uint8Array | Buffer
  width: number
  height: number
}

export interface PageGrid {
  colBounds: number[]
  rowBounds: number[]
  darkThreshold: number
  /** Which source each part of the grid came from — surfaced for diagnostics. */
  colSource: 'reference' | 'detected' | 'structural'
  rowSource: 'reference' | 'detected'
}

/**
 * Otsu's method: the grayscale cutoff that best splits a photo's own pixels
 * into two classes (ink vs. paper). A fixed cutoff tuned on a clean scan
 * (paper near-white) misreads a real phone photo, where paper background
 * commonly sits at 130-160 from lighting/shadow/paper tone — collapsing
 * ruled-line detection into one giant "dark" blob instead of the real lines.
 */
export function otsuThreshold(img: GrayImage): number {
  const hist = new Array(256).fill(0)
  for (let i = 0; i < img.data.length; i++) hist[img.data[i]]++
  const total = img.data.length
  let sum = 0
  for (let v = 0; v < 256; v++) sum += v * hist[v]
  let sumB = 0
  let weightB = 0
  let maxVariance = 0
  let threshold = 128
  for (let t = 0; t < 256; t++) {
    weightB += hist[t]
    if (weightB === 0) continue
    const weightF = total - weightB
    if (weightF === 0) break
    sumB += t * hist[t]
    const meanB = sumB / weightB
    const meanF = (sum - sumB) / weightF
    const variance = weightB * weightF * (meanB - meanF) * (meanB - meanF)
    if (variance > maxVariance) {
      maxVariance = variance
      threshold = t
    }
  }
  return threshold
}

export function darkColProfile(img: GrayImage, yFrom: number, yTo: number, darkThreshold: number): Float64Array {
  const { data, width } = img
  const profile = new Float64Array(width)
  const rows = yTo - yFrom
  for (let x = 0; x < width; x++) {
    let dark = 0
    for (let y = yFrom; y < yTo; y++) {
      if (data[y * width + x] < darkThreshold) dark++
    }
    profile[x] = dark / rows
  }
  return profile
}

export function darkRowProfile(img: GrayImage, xFrom: number, xTo: number, darkThreshold: number): Float64Array {
  const { data, width, height } = img
  const profile = new Float64Array(height)
  const cols = xTo - xFrom
  for (let y = 0; y < height; y++) {
    let dark = 0
    for (let x = xFrom; x < xTo; x++) {
      if (data[y * width + x] < darkThreshold) dark++
    }
    profile[y] = dark / cols
  }
  return profile
}

export function clusterPeaks(profile: Float64Array, threshold: number, minGap: number): number[] {
  const peaks: number[] = []
  let i = 0
  while (i < profile.length) {
    if (profile[i] > threshold) {
      let j = i
      let wsum = 0
      let sum = 0
      while (j < profile.length && profile[j] > threshold * 0.6) {
        wsum += profile[j] * j
        sum += profile[j]
        j++
      }
      peaks.push(Math.round(wsum / sum))
      i = j + minGap
    } else {
      i++
    }
  }
  return peaks
}

/** True when `candidate` boundaries are each within `tolerance`px of `ref`, scaled by `scale`. */
export function closeToReference(candidate: number[], ref: number[], scale: number, tolerance: number): boolean {
  if (candidate.length !== ref.length) return false
  return candidate.every((v, i) => Math.abs(v - ref[i] * scale) <= tolerance)
}

export function detectGrid(img: GrayImage): PageGrid {
  const scale = img.width / REF_WIDTH
  const refCols = REF_COLUMN_BOUNDS.map((x) => x * scale)
  const darkThreshold = otsuThreshold(img)

  let colBounds = refCols
  let colSource: PageGrid['colSource'] = 'reference'
  const colProfile = darkColProfile(img, Math.round(img.height * 0.15), Math.round(img.height * 0.95), darkThreshold)
  const detectedCols = clusterPeaks(colProfile, 0.25, Math.round(5 * scale))
  if (closeToReference(detectedCols, REF_COLUMN_BOUNDS, scale, 20 * scale)) {
    colBounds = detectedCols
    colSource = 'detected'
  } else {
    // A real photo's true column positions rarely land within tolerance of
    // the reference scaled linearly (perspective/crop differences distort
    // proportions unevenly across the table). No/L/B/D are printed as 4
    // equal-width columns on this form, so the correct 8-of-N boundaries
    // among the detected peaks (which can include extra noise peaks from
    // handwriting strokes lining up across rows) are the ones whose
    // No/L/B/D widths agree with each other.
    const structural = selectColumnBoundsByEqualWidth(detectedCols)
    if (structural !== undefined) {
      colBounds = structural
      colSource = 'structural'
    }
  }

  // Row lines are detected from the Date column specifically, not the full
  // table width: that column is normally blank/lightly filled, so its dark
  // profile reflects the ruled lines themselves rather than handwriting ink
  // density (which swamps a full-width profile on busy real photos).
  const dateColFrom = colBounds[0]
  const dateColTo = colBounds[1]
  const rowProfile = darkRowProfile(img, Math.round(dateColFrom + 5), Math.round(dateColTo - 5), darkThreshold)
  const detectedRows = clusterPeaks(rowProfile, 0.3, Math.round(15 * scale))

  let rowTop = REF_ROW_TOP * scale
  let pitch = REF_ROW_PITCH * scale
  let rowSource: PageGrid['rowSource'] = 'reference'
  let pitchConfirmed = false
  // Wide window: the true pitch varies with how close the photo was taken,
  // and the dominant-cluster estimate inside estimateRowPitchConfidence is
  // what separates real row spacing from header spacing, not this range.
  const pitchEstimate = estimateRowPitchConfidence(detectedRows, { minGap: 25 * scale, maxGap: 120 * scale })
  // Trust the detected pitch either when it matches the fixed-form reference
  // (typical case) or when it's internally self-consistent across several
  // gaps (a photo framed/cropped differently than the calibration reference
  // has its own valid, but different, pitch — reference match alone would
  // wrongly reject a clean detection just because it disagrees with a
  // constant derived from a different photo's proportions).
  const matchesReference = pitchEstimate !== undefined && Math.abs(pitchEstimate.pitch - pitch) <= 6 * scale
  const selfConsistent = pitchEstimate !== undefined && pitchEstimate.sampleCount >= 3 && pitchEstimate.maxDeviation <= 5 * scale
  if (pitchEstimate !== undefined && (matchesReference || selfConsistent)) {
    pitch = pitchEstimate.pitch
    pitchConfirmed = true
    const detectedTop = findGridStart(detectedRows, pitch, 6 * scale)
    if (detectedTop !== undefined) {
      rowTop = detectedTop
      rowSource = 'detected'
    }
  }

  // Bottom limit: prefer a detected line that's clearly beyond the last
  // regular-pitch row (the table's closing border), else fall back to the
  // image height so the reference pitch alone still produces a full page of
  // rows. Only trust a detected line here once pitch itself was confirmed —
  // otherwise detectedRows is noise (handwriting ink density, not ruled
  // lines) and its last entry truncates the grid to a few rows.
  const bottomCandidate = pitchConfirmed && detectedRows.length > 0 ? detectedRows[detectedRows.length - 1] : undefined
  const bottomLimit = bottomCandidate !== undefined && bottomCandidate > rowTop + pitch ? bottomCandidate : img.height - 10 * scale

  const rowBounds = pitchConfirmed
    ? generateRowBoundsSnapped(detectedRows, rowTop, pitch, bottomLimit)
    : generateRowBounds(rowTop, pitch, bottomLimit)
  return { colBounds, rowBounds, darkThreshold, colSource, rowSource }
}

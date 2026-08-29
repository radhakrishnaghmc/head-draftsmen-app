import { recognizeImages } from './ocr'
import { estimateRowPitch, findGridStart, generateRowBounds, buildMbMeasurementRows } from '../core/mbMeasurementExtract'
import type { MbMeasurementRow } from '../core/mbMeasurementExtract'

// Calibrated once from a clean sample page of the "L.F. No. 83" MB form (see
// scratchpad grid-prototype). This is a FIXED preprinted form — the same
// stationery on every page of every MB — so these are used as the primary
// source of truth for column bounds and row pitch, not just a fallback:
// per-page pixel detection is noisy (handwriting overlapping ruled lines,
// perspective/lighting differences between photos of different pages) and is
// only trusted as a refinement when it lands close to this reference.
const REF_WIDTH = 1275
const REF_COLUMN_BOUNDS = [78, 231, 697, 796, 895, 994, 1095, 1234] // 8 boundaries -> 7 columns
const REF_ROW_TOP = 258
const REF_ROW_PITCH = 48.71

const DARK_THRESHOLD = 150
const CELL_UPSCALE = 3
// A crop with less than this fraction of dark pixels is treated as blank —
// skips OCR entirely (faster, and avoids the model hallucinating text on
// blank cells).
const BLANK_CELL_DARK_FRACTION = 0.01

interface GrayImage {
  data: Buffer
  width: number
  height: number
}

async function toGrayscale(imageBuffer: Buffer): Promise<GrayImage> {
  // Loaded lazily so a missing/incompatible platform binary only breaks this
  // feature at point of use, matching the pattern in electron/ocr.ts.
  const sharp = (await import('sharp')).default
  const { data, info } = await sharp(imageBuffer).grayscale().raw().toBuffer({ resolveWithObject: true })
  return { data, width: info.width, height: info.height }
}

function darkColProfile(img: GrayImage, yFrom: number, yTo: number): Float64Array {
  const { data, width } = img
  const profile = new Float64Array(width)
  const rows = yTo - yFrom
  for (let x = 0; x < width; x++) {
    let dark = 0
    for (let y = yFrom; y < yTo; y++) {
      if (data[y * width + x] < DARK_THRESHOLD) dark++
    }
    profile[x] = dark / rows
  }
  return profile
}

function darkRowProfile(img: GrayImage, xFrom: number, xTo: number): Float64Array {
  const { data, width, height } = img
  const profile = new Float64Array(height)
  const cols = xTo - xFrom
  for (let y = 0; y < height; y++) {
    let dark = 0
    for (let x = xFrom; x < xTo; x++) {
      if (data[y * width + x] < DARK_THRESHOLD) dark++
    }
    profile[y] = dark / cols
  }
  return profile
}

function clusterPeaks(profile: Float64Array, threshold: number, minGap: number): number[] {
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

interface PageGrid {
  colBounds: number[]
  rowBounds: number[]
}

/** True when `candidate` boundaries are each within `tolerance`px of `ref`, scaled by `scale`. */
function closeToReference(candidate: number[], ref: number[], scale: number, tolerance: number): boolean {
  if (candidate.length !== ref.length) return false
  return candidate.every((v, i) => Math.abs(v - ref[i] * scale) <= tolerance)
}

function detectGrid(img: GrayImage): PageGrid {
  const scale = img.width / REF_WIDTH
  const refCols = REF_COLUMN_BOUNDS.map((x) => x * scale)

  let colBounds = refCols
  const colProfile = darkColProfile(img, Math.round(img.height * 0.15), Math.round(img.height * 0.95))
  const detectedCols = clusterPeaks(colProfile, 0.25, Math.round(5 * scale))
  if (closeToReference(detectedCols, REF_COLUMN_BOUNDS, scale, 20 * scale)) {
    colBounds = detectedCols
  }

  const left = colBounds[0]
  const right = colBounds[colBounds.length - 1]
  const rowProfile = darkRowProfile(img, Math.round(left + 5), Math.round(right - 5))
  const detectedRows = clusterPeaks(rowProfile, 0.28, 4)

  let rowTop = REF_ROW_TOP * scale
  let pitch = REF_ROW_PITCH * scale
  const detectedPitch = estimateRowPitch(detectedRows, { minGap: 30 * scale, maxGap: 70 * scale })
  if (detectedPitch !== undefined && Math.abs(detectedPitch - pitch) <= 6 * scale) {
    pitch = detectedPitch
    const detectedTop = findGridStart(detectedRows, pitch, 6 * scale)
    if (detectedTop !== undefined) rowTop = detectedTop
  }

  // Bottom limit: prefer a detected line that's clearly beyond the last
  // regular-pitch row (the table's closing border), else fall back to the
  // image height so the reference pitch alone still produces a full page of
  // rows.
  const bottomCandidate = detectedRows.length > 0 ? detectedRows[detectedRows.length - 1] : undefined
  const bottomLimit = bottomCandidate !== undefined && bottomCandidate > rowTop + pitch ? bottomCandidate : img.height - 10 * scale

  const rowBounds = generateRowBounds(rowTop, pitch, bottomLimit)
  return { colBounds, rowBounds }
}

async function cropCell(
  sharpInstance: import('sharp').Sharp,
  x0: number,
  x1: number,
  y0: number,
  y1: number,
  imgWidth: number,
  imgHeight: number
): Promise<{ buffer: Buffer; darkFraction: number } | undefined> {
  const left = Math.max(0, Math.round(x0) + 2)
  const top = Math.max(0, Math.round(y0) + 2)
  // A page's real dimensions can differ slightly from the calibration
  // reference (different scanner crop/aspect ratio), so the generated grid
  // bounds — especially the last row/column, which is allowed to overshoot
  // to capture a partial band — can land past the actual image edge. Clamp
  // rather than let sharp's extract() throw, and skip a cell that's
  // entirely off-page.
  if (left >= imgWidth || top >= imgHeight) return undefined
  const width = Math.min(imgWidth - left, Math.max(1, Math.round(x1) - Math.round(x0) - 4))
  const height = Math.min(imgHeight - top, Math.max(1, Math.round(y1) - Math.round(y0) - 4))

  const sharp = (await import('sharp')).default
  const { data, info } = await sharpInstance
    .clone()
    .extract({ left, top, width, height })
    .grayscale()
    .raw()
    .toBuffer({ resolveWithObject: true })

  let dark = 0
  for (let i = 0; i < data.length; i++) if (data[i] < DARK_THRESHOLD) dark++
  const darkFraction = dark / data.length
  if (darkFraction < BLANK_CELL_DARK_FRACTION) return { buffer: Buffer.alloc(0), darkFraction }

  const buffer = await sharp(data, { raw: { width: info.width, height: info.height, channels: 1 } })
    .resize({ width: info.width * CELL_UPSCALE })
    .png()
    .toBuffer()
  return { buffer, darkFraction }
}

/** Extract the MB measurement grid (as raw string cells, 7 columns) from one page image. */
async function extractPageGrid(imageBuffer: Buffer, onProgress?: (done: number, total: number) => void): Promise<string[][]> {
  const sharp = (await import('sharp')).default
  const gray = await toGrayscale(imageBuffer)
  const { colBounds, rowBounds } = detectGrid(gray)
  const sharpInstance = sharp(imageBuffer)

  interface CellSlot {
    row: number
    col: number
  }
  const cellBuffers: Buffer[] = []
  const slots: CellSlot[] = []
  const grid: string[][] = Array.from({ length: rowBounds.length - 1 }, () => Array(colBounds.length - 1).fill(''))

  for (let row = 0; row < rowBounds.length - 1; row++) {
    for (let col = 0; col < colBounds.length - 1; col++) {
      const cropped = await cropCell(
        sharpInstance,
        colBounds[col],
        colBounds[col + 1],
        rowBounds[row],
        rowBounds[row + 1],
        gray.width,
        gray.height
      )
      if (!cropped || cropped.buffer.length === 0) continue
      cellBuffers.push(cropped.buffer)
      slots.push({ row, col })
    }
  }

  let done = 0
  onProgress?.(0, cellBuffers.length)
  const results = await recognizeImages(cellBuffers, () => onProgress?.(++done, cellBuffers.length))
  for (let i = 0; i < slots.length; i++) {
    const { row, col } = slots[i]
    const text = results[i]
      .slice()
      .sort((a, b) => a.top - b.top)
      .map((l) => l.text)
      .join('\n')
      .trim()
    grid[row][col] = text
  }

  return grid
}

export interface MbMeasurementOcrProgress {
  page: number
  totalPages: number
  cellsDone: number
  cellsTotal: number
}

/** Extract MB measurement rows from one or more page images (a whole scanned MB card/book), in page order. */
export async function extractMbMeasurementSheet(
  pageBuffers: Buffer[],
  onProgress?: (progress: MbMeasurementOcrProgress) => void
): Promise<MbMeasurementRow[]> {
  const rows: MbMeasurementRow[] = []
  for (let i = 0; i < pageBuffers.length; i++) {
    const grid = await extractPageGrid(pageBuffers[i], (cellsDone, cellsTotal) => {
      onProgress?.({ page: i + 1, totalPages: pageBuffers.length, cellsDone, cellsTotal })
    })
    rows.push(...buildMbMeasurementRows(grid))
  }
  return rows
}

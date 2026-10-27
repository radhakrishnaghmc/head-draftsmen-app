import { recognizeImage, recognizeImages } from './ocr'
import { buildMbMeasurementRows } from '../core/mbMeasurementExtract'
import type { MbMeasurementRow } from '../core/mbMeasurementExtract'
import { detectGrid } from '../core/mbGridDetect'
import { detectGridAuto } from '../core/mbGridAuto'
import type { GrayImage } from '../core/mbGridDetect'

const CELL_UPSCALE = 3
// Crop padding as a fraction of the cell's own width/height (see cropCell).
const CELL_PAD_X = 0.12
const CELL_PAD_Y = 0.18
// White margin added around the upscaled crop before recognition (px).
const CELL_MARGIN = 40
// A crop with less than this fraction of dark pixels is treated as blank —
// skips OCR entirely (faster, and avoids the model hallucinating text on
// blank cells).
const BLANK_CELL_DARK_FRACTION = 0.01

async function toGrayscale(imageBuffer: Buffer): Promise<GrayImage> {
  // Loaded lazily so a missing/incompatible platform binary only breaks this
  // feature at point of use, matching the pattern in electron/ocr.ts.
  const sharp = (await import('sharp')).default
  const { data, info } = await sharp(imageBuffer).grayscale().raw().toBuffer({ resolveWithObject: true })
  return { data, width: info.width, height: info.height }
}

export async function cropCell(
  sharpInstance: import('sharp').Sharp,
  x0: number,
  x1: number,
  y0: number,
  y1: number,
  imgWidth: number,
  imgHeight: number,
  darkThreshold: number
): Promise<{ buffer: Buffer; darkFraction: number } | undefined> {
  // Handwriting does not respect the ruled cell it belongs to: figures sit
  // high against the line above, and an averaged entry's numerator
  // ("6.70+6.50" over "2") regularly starts left of its column divider. A
  // crop cut exactly on the rules therefore clips real ink, and the model
  // then reads "6.10+6.20" as "1046-20". Padding the crop keeps the whole
  // figure — the neighbouring columns' own figures stay clear of the pad
  // because each is written well inside its (much wider) column.
  const cellWidth = Math.round(x1) - Math.round(x0)
  const cellHeight = Math.round(y1) - Math.round(y0)
  const padX = Math.round(cellWidth * CELL_PAD_X)
  const padY = Math.round(cellHeight * CELL_PAD_Y)
  const left = Math.max(0, Math.round(x0) + 2 - padX)
  const top = Math.max(0, Math.round(y0) + 2 - padY)
  // A page's real dimensions can differ slightly from the calibration
  // reference (different scanner crop/aspect ratio), so the generated grid
  // bounds — especially the last row/column, which is allowed to overshoot
  // to capture a partial band — can land past the actual image edge. Clamp
  // rather than let sharp's extract() throw, and skip a cell that's
  // entirely off-page.
  if (left >= imgWidth || top >= imgHeight) return undefined
  const width = Math.min(imgWidth - left, Math.max(1, cellWidth - 4 + padX * 2))
  const height = Math.min(imgHeight - top, Math.max(1, cellHeight - 4 + padY * 2))

  const sharp = (await import('sharp')).default
  const crop = sharpInstance.clone().extract({ left, top, width, height }).grayscale()
  // Blankness is judged on the untouched crop, against the whole page's ink
  // threshold: contrast-stretching a blank cell first would amplify paper
  // grain into "ink" and put the cell back in front of the model.
  const { data, info } = await crop.clone().raw().toBuffer({ resolveWithObject: true })
  // Only the cell interior counts: the padded crop also holds the ruled lines
  // around the cell, which would make every empty cell look inked.
  const insetX = padX + 4
  const insetY = padY + 4
  let dark = 0
  let total = 0
  for (let y = insetY; y < info.height - insetY; y++) {
    for (let x = insetX; x < info.width - insetX; x++) {
      total++
      if (data[y * info.width + x] < darkThreshold) dark++
    }
  }
  const darkFraction = total > 0 ? dark / total : 0
  if (darkFraction < BLANK_CELL_DARK_FRACTION) return { buffer: Buffer.alloc(0), darkFraction }

  // Phone photos of a book page carry uneven lighting and yellowed paper;
  // stretching the crop's own contrast lifts pale ink off the background
  // before the model sees it.
  // The white margin matters: the detector inside the model finds nothing in a
  // crop whose text runs to the edge, and a cell cut on the ruled lines always
  // does. Measured on real sample pages, adding it alone recovered figures
  // that came back empty without it.
  const buffer = await crop
    .clone()
    .normalise()
    .resize({ width: info.width * CELL_UPSCALE })
    .extend({ top: CELL_MARGIN, bottom: CELL_MARGIN, left: CELL_MARGIN, right: CELL_MARGIN, background: '#ffffff' })
    .png()
    .toBuffer()
  return { buffer, darkFraction }
}

/**
 * Ink mask: a pixel is ink when it is clearly darker than its local
 * surroundings. Faint ruled lines and uneven phone/scanner lighting survive
 * this where a single global threshold loses them.
 */
async function inkMask(imageBuffer: Buffer, delta = 25): Promise<GrayImage> {
  const sharp = (await import('sharp')).default
  const g = await sharp(imageBuffer).grayscale().raw().toBuffer({ resolveWithObject: true })
  const bg = await sharp(imageBuffer).grayscale().blur(18).raw().toBuffer()
  const bin = Buffer.alloc(g.data.length, 255)
  for (let i = 0; i < bin.length; i++) if (bg[i] - g.data[i] > delta) bin[i] = 0
  return { data: bin, width: g.info.width, height: g.info.height }
}

// The printed header words: how many show up decides which way up a page is.
async function headerWordCount(imageBuffer: Buffer): Promise<number> {
  const lines = await recognizeImage(imageBuffer)
  const text = lines.map((l) => l.text).join(' ')
  return (text.match(/measurement|description|contents|masure|work ect/gi) ?? []).length
}

interface OrientedPage {
  buffer: Buffer
  grid: { colBounds: number[]; rowBounds: number[] }
}

/**
 * Find which of the four orientations the page is in (books get scanned
 * sideways), returning the buffer turned upright plus its grid. An upright
 * page costs one extra whole-page read of the header.
 */
export async function orientPage(imageBuffer: Buffer): Promise<OrientedPage | undefined> {
  const sharp = (await import('sharp')).default
  let best: { page: OrientedPage; words: number } | undefined
  for (const deg of [0, 270, 90, 180]) {
    const buffer = deg === 0 ? imageBuffer : await sharp(imageBuffer).rotate(deg).png().toBuffer()
    const grid = detectGridAuto(await inkMask(buffer))
    if (!grid) continue
    const words = await headerWordCount(buffer)
    const page = { buffer, grid }
    if (words >= 2) return page
    if (!best || words > best.words) best = { page, words }
  }
  return best?.page
}

/** Extract the MB measurement grid (as raw string cells, 7 columns) from one page image. */
async function extractPageGrid(imageBuffer: Buffer, onProgress?: (done: number, total: number) => void): Promise<string[][]> {
  const sharp = (await import('sharp')).default
  const oriented = await orientPage(imageBuffer)
  if (oriented) imageBuffer = oriented.buffer
  const gray = await toGrayscale(imageBuffer)
  const fallback = oriented ? undefined : detectGrid(gray)
  const { colBounds, rowBounds } = oriented?.grid ?? fallback!
  const darkThreshold = fallback?.darkThreshold ?? detectGrid(gray).darkThreshold
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
        gray.height,
        darkThreshold
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

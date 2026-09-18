import { recognizeImages } from './ocr'
import { buildMbMeasurementRows } from '../core/mbMeasurementExtract'
import type { MbMeasurementRow } from '../core/mbMeasurementExtract'
import { detectGrid } from '../core/mbGridDetect'
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

async function cropCell(
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
  let dark = 0
  for (let i = 0; i < data.length; i++) if (data[i] < darkThreshold) dark++
  const darkFraction = dark / data.length
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

/** Extract the MB measurement grid (as raw string cells, 7 columns) from one page image. */
async function extractPageGrid(imageBuffer: Buffer, onProgress?: (done: number, total: number) => void): Promise<string[][]> {
  const sharp = (await import('sharp')).default
  const gray = await toGrayscale(imageBuffer)
  const { colBounds, rowBounds, darkThreshold } = detectGrid(gray)
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

import { parentPort } from 'worker_threads'
import type { Worker as TesseractWorker } from 'tesseract.js'
import type { OcrWord } from '../core/ocrTableReconstruct'

// Runs on its own OS thread, same reasoning as embeddingsWorker.ts — OCR on
// a full-page photo can take a few seconds, and doing that on Electron's
// single-threaded main process would freeze the entire app (every IPC
// call, window repaint) for the duration.

let workerPromise: Promise<TesseractWorker> | null = null

async function getWorker(langPath: string, cachePath: string): Promise<TesseractWorker> {
  if (!workerPromise) {
    workerPromise = (async () => {
      const { createWorker } = await import('tesseract.js')
      // gzip: true because the bundled traineddata ships gzipped (see
      // resources/tessdata/) — smaller to bundle, and Tesseract.js
      // decompresses it itself. langPath (read the bundled .gz from) and
      // cachePath (write the decompressed copy to) are deliberately
      // different: langPath lives inside the packaged app's resources,
      // which isn't guaranteed writable, so the decompressed cache goes to
      // userData instead (see electron/ocr.ts's ocrCacheDir()).
      return createWorker('eng', 1, { langPath, cachePath, gzip: true })
    })()
  }
  return workerPromise
}

interface Request {
  id: number
  imageBuffer: Buffer
  langPath: string
  cachePath: string
}

// Real phone photos of paper documents are usually under-sharpened, unevenly
// lit, and small relative to the page's text size — all of which measurably
// hurt Tesseract's accuracy. Grayscale + contrast normalization + sharpening
// + upscaling (sharp is already bundled transitively via
// @huggingface/transformers, so this adds no new native dependency to
// package) consistently cleans up the recognized text without needing any
// perspective/deskew correction. Best-effort: if preprocessing itself fails
// for any reason, fall back to the original image rather than losing the OCR
// attempt entirely.
async function preprocess(imageBuffer: Buffer): Promise<Buffer> {
  try {
    const sharp = (await import('sharp')).default
    const meta = await sharp(imageBuffer).metadata()
    const width = meta.width ?? 0
    return await sharp(imageBuffer)
      .rotate()
      .grayscale()
      .normalize()
      .sharpen()
      .resize(width && width < 2400 ? { width: 2400 } : undefined)
      .toBuffer()
  } catch {
    return imageBuffer
  }
}

parentPort?.on('message', async (msg: Request) => {
  try {
    const worker = await getWorker(msg.langPath, msg.cachePath)
    const image = await preprocess(msg.imageBuffer)
    const result = await worker.recognize(image, {}, { blocks: true })
    const words: OcrWord[] = []
    for (const block of result.data.blocks ?? []) {
      for (const para of block.paragraphs ?? []) {
        for (const line of para.lines ?? []) {
          for (const w of line.words ?? []) {
            words.push({ text: w.text, bbox: w.bbox })
          }
        }
      }
    }
    parentPort?.postMessage({ id: msg.id, words })
  } catch (e) {
    parentPort?.postMessage({ id: msg.id, error: e instanceof Error ? e.message : String(e) })
  }
})

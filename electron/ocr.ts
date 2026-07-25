import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { Worker } from 'worker_threads'
import { app } from 'electron'
import sharp from 'sharp'

export interface OcrLine {
  /** The recognized text of one detected line, in no particular order (callers sort by `top`). */
  text: string
  top: number
}

// Bundled OCR model dir candidates, matching the existing bundled-resource
// pattern (see modelDir() in electron/embeddings.ts) — first for dev, then
// packaged-app layouts. Read-only in a packaged app; only ever read from.
function ocrModelDir(): string | undefined {
  const candidates = [
    path.join(process.resourcesPath, 'ocr-models'),
    path.join(app.getAppPath(), 'resources', 'ocr-models'),
    path.join(app.getAppPath(), '..', 'resources', 'ocr-models')
  ]
  return candidates.find((p) => {
    try {
      fs.accessSync(p)
      return true
    } catch {
      return false
    }
  })
}

interface Pending {
  resolve: (lines: OcrLine[]) => void
  reject: (err: Error) => void
}

// The OCR engine runs in a worker thread, not on this (the main) process's
// event loop — same reasoning as electron/embeddings.ts. Spawned once and
// reused across photos/calls.
let worker: Worker | null = null
let nextId = 0
const pending = new Map<number, Pending>()

function getWorker(): Worker {
  if (worker) return worker
  const w = new Worker(path.join(__dirname, 'ocrWorker.js'))
  w.on('message', (msg: { id: number; lines?: OcrLine[]; error?: string }) => {
    const p = pending.get(msg.id)
    if (!p) return
    pending.delete(msg.id)
    if (msg.error) p.reject(new Error(msg.error))
    else p.resolve(msg.lines ?? [])
  })
  w.on('error', (err) => {
    for (const p of pending.values()) p.reject(err)
    pending.clear()
    worker = null // let the next call spawn a fresh worker
  })
  worker = w
  return w
}

// A full-resolution phone photo (routinely 3000-4000px on the long side) fed
// straight to @gutenye/ocr-node's onnxruntime-node inference crashes the
// whole app with a native SIGTRAP deep inside onnxruntime's memory arena —
// reproduced directly against a real photo through the actual compiled OCR
// pipeline (not a guess): the identical call succeeds every time once the
// image is capped to this size first, with no loss in recognized text (the
// same real photo read the same 53 lines either way) since paper-estimate
// text never needs anywhere near full phone-camera resolution to read.
const MAX_IMAGE_DIMENSION = 1800

async function capImageSize(imageBuffer: Buffer): Promise<Buffer> {
  return sharp(imageBuffer)
    .rotate() // apply EXIF orientation before resizing, so a sideways photo is capped on the right axis
    .resize({ width: MAX_IMAGE_DIMENSION, height: MAX_IMAGE_DIMENSION, fit: 'inside', withoutEnlargement: true })
    .toBuffer()
}

/**
 * Run OCR on one image (a photo of a paper estimate page) and return its
 * recognized lines with their vertical position, in no particular order.
 *
 * @gutenye/ocr-node's detect() only accepts a file path, not a Buffer, so
 * the image is written to a temp file first (cleaned up after use).
 */
export async function recognizeImage(imageBuffer: Buffer): Promise<OcrLine[]> {
  const dir = ocrModelDir()
  if (!dir) throw new Error('Bundled OCR model is missing from the app.')

  const capped = await capImageSize(imageBuffer)
  const tempPath = path.join(os.tmpdir(), `hda-ocr-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.png`)
  fs.writeFileSync(tempPath, capped)
  try {
    const id = nextId++
    const w = getWorker()
    return await new Promise((resolve, reject) => {
      pending.set(id, { resolve, reject })
      w.postMessage({ id, imagePath: tempPath, modelDir: dir })
    })
  } finally {
    fs.unlink(tempPath, () => {})
  }
}

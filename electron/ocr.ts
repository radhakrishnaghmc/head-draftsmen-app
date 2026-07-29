import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { fork, type ChildProcess } from 'child_process'
import { app } from 'electron'

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

// Thrown when the OCR child process dies mid-inference (e.g. the native
// SIGTRAP described below), as opposed to returning a normal error. It's the
// signal recognizeImage() uses to retry the page at a smaller resolution.
class OcrProcessCrashError extends Error {}

// OCR runs in an isolated child PROCESS (child_process.fork), not a worker
// thread. A native onnxruntime crash on a large image cannot be caught in JS,
// so if it were a worker thread it would take the whole app down with it. In a
// separate process the fault only kills the child, and the parent retries the
// page at a smaller, proven-safe size (see recognizeImage). ELECTRON_RUN_AS_NODE
// makes fork() launch the Electron binary as plain Node so onnxruntime-node's
// N-API addon loads exactly as it does today. Spawned once and reused across
// pages; respawned after a crash.
let child: ChildProcess | null = null
let nextId = 0
const pending = new Map<number, Pending>()

function getChild(): ChildProcess {
  if (child) return child
  const c = fork(path.join(__dirname, 'ocrWorker.js'), [], {
    env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }
  })
  c.on('message', (msg: { id: number; lines?: OcrLine[]; error?: string }) => {
    const p = pending.get(msg.id)
    if (!p) return
    pending.delete(msg.id)
    if (msg.error) p.reject(new Error(msg.error))
    else p.resolve(msg.lines ?? [])
  })
  const onDead = (info: string): void => {
    // The child died with requests in flight — fail them with the crash marker
    // so recognizeImage can retry smaller, and drop the handle so the next call
    // spawns a fresh child.
    for (const p of pending.values()) p.reject(new OcrProcessCrashError(info))
    pending.clear()
    if (child === c) child = null
  }
  c.on('exit', (code, signal) => onDead(`OCR process exited (code=${code}, signal=${signal})`))
  c.on('error', (err) => onDead(err.message))
  child = c
  return c
}

// The higher-resolution pass: a phone photo of a dense paper estimate needs
// real resolution for small print to survive recognition, so the primary pass
// caps the long side here rather than at the historical safe-but-lossy 1800.
const PRIMARY_MAX_DIMENSION = 3000
// The fallback pass, only used if the primary pass crashes the OCR process: the
// long-proven size that never triggered the native onnxruntime fault.
const FALLBACK_MAX_DIMENSION = 1800

async function capImageSize(imageBuffer: Buffer, maxDimension: number): Promise<Buffer> {
  // Loaded lazily (not a top-level import) so that a missing/incompatible
  // platform binary can only ever break the OCR feature at the point of use,
  // never crash the whole app at startup when this module is first required.
  const sharp = (await import('sharp')).default
  return sharp(imageBuffer)
    .rotate() // apply EXIF orientation before resizing, so a sideways photo is capped on the right axis
    .resize({ width: maxDimension, height: maxDimension, fit: 'inside', withoutEnlargement: true })
    .toBuffer()
}

async function runOcrPass(imageBuffer: Buffer, dir: string, maxDimension: number): Promise<OcrLine[]> {
  const capped = await capImageSize(imageBuffer, maxDimension)
  const tempPath = path.join(os.tmpdir(), `hda-ocr-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.png`)
  fs.writeFileSync(tempPath, capped)
  try {
    const id = nextId++
    const c = getChild()
    return await new Promise<OcrLine[]>((resolve, reject) => {
      pending.set(id, { resolve, reject })
      c.send({ id, imagePath: tempPath, modelDir: dir }, (err) => {
        // The channel closed between spawning and sending (child already dead):
        // surface it as a crash so the caller can fall back.
        if (err && pending.delete(id)) reject(new OcrProcessCrashError(err.message))
      })
    })
  } finally {
    fs.unlink(tempPath, () => {})
  }
}

/**
 * Run OCR on one image (a photo of a paper estimate page) and return its
 * recognized lines with their vertical position, in no particular order.
 *
 * Runs a higher-resolution pass first for accuracy; if that crashes the
 * isolated OCR process, retries this page once at the smaller safe size.
 *
 * @gutenye/ocr-node's detect() only accepts a file path, not a Buffer, so
 * the image is written to a temp file first (cleaned up after use).
 */
export async function recognizeImage(imageBuffer: Buffer): Promise<OcrLine[]> {
  const dir = ocrModelDir()
  if (!dir) throw new Error('Bundled OCR model is missing from the app.')

  try {
    return await runOcrPass(imageBuffer, dir, PRIMARY_MAX_DIMENSION)
  } catch (e) {
    if (e instanceof OcrProcessCrashError) {
      return await runOcrPass(imageBuffer, dir, FALLBACK_MAX_DIMENSION)
    }
    throw e
  }
}

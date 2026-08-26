import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { fork, type ChildProcess } from 'child_process'
import { app } from 'electron'

export interface OcrLine {
  /** The recognized text of one detected line, in no particular order (callers sort by `top`). */
  text: string
  top: number
  /** Bounding box of the detected line in the OCR-input image's pixels: [x, y, width, height]. Enables image-based table reconstruction (offline, no extra model). */
  box?: [number, number, number, number]
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

// OCR runs in isolated child PROCESSES (child_process.fork), not worker
// threads. A native onnxruntime crash on a large image cannot be caught in JS,
// so in a worker thread it would take the whole app down with it. In a separate
// process the fault only kills that child, and the parent retries the page at a
// smaller, proven-safe size (see recognizeImage). ELECTRON_RUN_AS_NODE makes
// fork() launch the Electron binary as plain Node so onnxruntime-node's N-API
// addon loads exactly as it does today.
//
// A *pool* of children (not one) so a batch — every photo needs several OCR
// passes — spreads across CPU cores instead of serialising on a single process:
// onnxruntime inference is single-threaded per call, so one child ≈ one core,
// and the pool is what turns a multi-photo batch from sequential into parallel.
interface OcrChild {
  proc: ChildProcess
  pending: Map<number, Pending>
}

// Leave a core free for the UI / main process; never fewer than 1, never more
// than 4 (each child loads its own copy of the model, so the pool costs RAM).
const POOL_SIZE = Math.max(1, Math.min(4, (os.cpus()?.length ?? 2) - 1))
let pool: OcrChild[] = []
let nextId = 0

function spawnChild(): OcrChild {
  const proc = fork(path.join(__dirname, 'ocrWorker.js'), [], {
    env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' }
  })
  const entry: OcrChild = { proc, pending: new Map() }
  proc.on('message', (msg: { id: number; lines?: OcrLine[]; error?: string }) => {
    const p = entry.pending.get(msg.id)
    if (!p) return
    entry.pending.delete(msg.id)
    if (msg.error) p.reject(new Error(msg.error))
    else p.resolve(msg.lines ?? [])
  })
  const onDead = (info: string): void => {
    // This child died with requests in flight — fail them with the crash marker
    // so recognizeImage can retry smaller, and drop it from the pool so the next
    // pick spawns a fresh one.
    for (const p of entry.pending.values()) p.reject(new OcrProcessCrashError(info))
    entry.pending.clear()
    pool = pool.filter((c) => c !== entry)
  }
  proc.on('exit', (code, signal) => onDead(`OCR process exited (code=${code}, signal=${signal})`))
  proc.on('error', (err) => onDead(err.message))
  pool.push(entry)
  return entry
}

// Reuse an idle child if there is one; otherwise grow the pool up to POOL_SIZE
// (so a single photo's passes never spawn more than one child and reload the
// model N times); otherwise queue on the least-busy child.
function pickChild(): OcrChild {
  const free = pool.find((c) => c.pending.size === 0)
  if (free) return free
  if (pool.length < POOL_SIZE) return spawnChild()
  return pool.reduce((a, b) => (b.pending.size < a.pending.size ? b : a))
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

// A native onnxruntime/detect() hang (as opposed to a crash) never fires
// 'exit'/'error'/a reply message at all — with no timeout, a real report:
// the upload button got stuck on "Reading…" forever, no error, nothing in
// the log, only recoverable by restarting the app. This bounds that: past
// OCR_TIMEOUT_MS the pending request is treated exactly like a process
// crash (same OcrProcessCrashError path recognizeImage already retries
// once at a smaller, safer resolution), and the wedged child is killed and
// dropped from the pool so it can't silently keep swallowing future pages.
const OCR_TIMEOUT_MS = 90_000

async function runOcrPass(imageBuffer: Buffer, dir: string, maxDimension: number): Promise<OcrLine[]> {
  const capped = await capImageSize(imageBuffer, maxDimension)
  const tempPath = path.join(os.tmpdir(), `hda-ocr-${Date.now()}-${Math.random().toString(36).slice(2, 8)}.png`)
  fs.writeFileSync(tempPath, capped)
  try {
    const id = nextId++
    const c = pickChild()
    return await new Promise<OcrLine[]>((resolve, reject) => {
      const timer = setTimeout(() => {
        if (!c.pending.delete(id)) return // already resolved/rejected — timer lost the race, harmless
        c.proc.kill()
        reject(new OcrProcessCrashError(`OCR timed out after ${OCR_TIMEOUT_MS / 1000}s (page likely wedged the OCR process).`))
      }, OCR_TIMEOUT_MS)
      c.pending.set(id, {
        resolve: (lines) => {
          clearTimeout(timer)
          resolve(lines)
        },
        reject: (err) => {
          clearTimeout(timer)
          reject(err)
        }
      })
      c.proc.send({ id, imagePath: tempPath, modelDir: dir }, (err) => {
        // The channel closed between spawning and sending (child already dead):
        // surface it as a crash so the caller can fall back.
        if (err && c.pending.delete(id)) {
          clearTimeout(timer)
          reject(new OcrProcessCrashError(err.message))
        }
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

/**
 * OCR many images and return their lines in the SAME order as the input.
 *
 * The win over a plain `for … await recognizeImage` is parallelism: it runs up
 * to POOL_SIZE pages at once — one per OCR child process, i.e. one per spare CPU
 * core — instead of one page at a time. Concurrency is capped at the pool size
 * (rather than firing every page at once) so a big multi-page PDF doesn't spawn
 * hundreds of temp images / model queue entries simultaneously; each worker
 * pulls the next page as it finishes, keeping every core busy end to end.
 */
export async function recognizeImages(imageBuffers: Buffer[]): Promise<OcrLine[][]> {
  const results: OcrLine[][] = new Array(imageBuffers.length)
  let next = 0
  const workerCount = Math.max(1, Math.min(POOL_SIZE, imageBuffers.length))
  const worker = async (): Promise<void> => {
    for (let i = next++; i < imageBuffers.length; i = next++) {
      results[i] = await recognizeImage(imageBuffers[i])
    }
  }
  await Promise.all(Array.from({ length: workerCount }, () => worker()))
  return results
}

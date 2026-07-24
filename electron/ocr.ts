import * as fs from 'fs'
import * as path from 'path'
import { Worker } from 'worker_threads'
import { app } from 'electron'
import type { OcrWord } from '../core/ocrTableReconstruct'

// Local traineddata dir candidates, matching the existing bundled-resource
// pattern (see modelDir() in electron/embeddings.ts). Read-only in a
// packaged app (inside the app bundle/install dir) — only ever read from,
// never written to.
function tessdataDir(): string | undefined {
  const candidates = [
    path.join(process.resourcesPath, 'tessdata'),
    path.join(app.getAppPath(), 'resources', 'tessdata'),
    path.join(app.getAppPath(), '..', 'resources', 'tessdata')
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

// Tesseract.js decompresses the bundled .traineddata.gz into a plain
// .traineddata file on first use and caches it for reuse — that cache
// write must land somewhere writable, which the bundled resources
// directory above is not guaranteed to be (e.g. inside a packaged macOS
// .app bundle). userData always is.
function ocrCacheDir(): string {
  const dir = path.join(app.getPath('userData'), 'ocr-cache')
  fs.mkdirSync(dir, { recursive: true })
  return dir
}

interface Pending {
  resolve: (words: OcrWord[]) => void
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
  w.on('message', (msg: { id: number; words?: OcrWord[]; error?: string }) => {
    const p = pending.get(msg.id)
    if (!p) return
    pending.delete(msg.id)
    if (msg.error) p.reject(new Error(msg.error))
    else p.resolve(msg.words ?? [])
  })
  w.on('error', (err) => {
    for (const p of pending.values()) p.reject(err)
    pending.clear()
    worker = null // let the next call spawn a fresh worker
  })
  worker = w
  return w
}

/** Run OCR on one image (a photo of a paper estimate page) and return its recognized words with positions. */
export async function recognizeImage(imageBuffer: Buffer): Promise<OcrWord[]> {
  const dir = tessdataDir()
  if (!dir) throw new Error('Bundled OCR language data is missing from the app.')
  const id = nextId++
  const w = getWorker()
  return new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject })
    w.postMessage({ id, imageBuffer, langPath: dir, cachePath: ocrCacheDir() })
  })
}

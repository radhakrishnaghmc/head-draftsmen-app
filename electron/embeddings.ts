import * as fs from 'fs'
import * as path from 'path'
import { Worker } from 'worker_threads'
import { app } from 'electron'

// Local model dir candidates, matching the existing bundled-resource pattern
// (see boqTemplateFile()/tenderNoticeTemplateFile() in main.ts) — first for
// dev, then packaged-app layouts.
function modelDir(): string | undefined {
  const candidates = [
    path.join(process.resourcesPath, 'models'),
    path.join(app.getAppPath(), 'resources', 'models'),
    path.join(app.getAppPath(), '..', 'resources', 'models')
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
  resolve: (vectors: number[][]) => void
  reject: (err: Error) => void
}

// The embedding model runs in a worker thread, not on this (the main)
// process's event loop — computing embeddings for a few hundred rate-
// database entries takes 15s+, and doing that synchronously in the main
// process would freeze the entire app (every IPC call, window repaint) for
// the whole duration, which can get the app force-quit by the OS as
// "not responding". The worker is spawned once and reused.
let worker: Worker | null = null
let nextId = 0
const pending = new Map<number, Pending>()

function getWorker(): Worker {
  if (worker) return worker
  const w = new Worker(path.join(__dirname, 'embeddingsWorker.js'))
  w.on('message', (msg: { id: number; vectors?: number[][]; error?: string }) => {
    const p = pending.get(msg.id)
    if (!p) return
    pending.delete(msg.id)
    if (msg.error) p.reject(new Error(msg.error))
    else p.resolve(msg.vectors ?? [])
  })
  w.on('error', (err) => {
    for (const p of pending.values()) p.reject(err)
    pending.clear()
    worker = null // let the next call spawn a fresh worker
  })
  worker = w
  return w
}

function embedBatch(texts: string[], dir: string): Promise<number[][]> {
  const id = nextId++
  const w = getWorker()
  return new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject })
    w.postMessage({ id, texts, modelDir: dir })
  })
}

// Sending the full rate-database (700-800+ real, ~200-char descriptions) as
// one giant batch to onnxruntime-node was observed to crash or hang the
// process — reproduced directly against the bundled model outside the app.
// Chunking into small batches sent sequentially avoids whatever internal
// limit that hits, at a small, easily-acceptable cost to total time.
const CHUNK_SIZE = 16

/** Compute normalized sentence embeddings for a batch of texts, off the main thread. */
export async function embedTexts(texts: string[]): Promise<number[][]> {
  if (texts.length === 0) return []
  const dir = modelDir()
  if (!dir) throw new Error('Bundled embedding model is missing from the app.')

  const vectors: number[][] = []
  for (let i = 0; i < texts.length; i += CHUNK_SIZE) {
    const chunk = texts.slice(i, i + CHUNK_SIZE)
    vectors.push(...(await embedBatch(chunk, dir)))
  }
  return vectors
}

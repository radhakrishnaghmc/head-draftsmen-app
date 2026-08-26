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

// Serializes every request sent to the worker, one at a time — two
// callers each awaiting their own embedTexts() (e.g. resolveForRow's
// label/column pair, or a background warm-up overlapping a real request)
// would otherwise post concurrent messages, and onnxruntime-node's own
// inference session is not safe to re-enter concurrently (see CHUNK_SIZE's
// doc comment above for another instability this same native binary hits
// under load). Chaining onto this tail costs nothing once the queue is
// empty and costs only wall-clock time (not correctness) when it isn't.
let requestQueue: Promise<unknown> = Promise.resolve()

function embedBatch(texts: string[], dir: string): Promise<number[][]> {
  const run = () => {
    const id = nextId++
    const w = getWorker()
    return new Promise<number[][]>((resolve, reject) => {
      pending.set(id, { resolve, reject })
      w.postMessage({ id, texts, modelDir: dir })
    })
  }
  const result = requestQueue.then(run, run)
  requestQueue = result.catch(() => {})
  return result
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

/**
 * Fires the ~15s+ model load in the background right after the app starts,
 * so it's already warm by the time a user's first placeholder-matching call
 * (e.g. a document Preview) needs it — without this, that first call is the
 * one left waiting on the load, which reads as the whole preview hanging.
 * Fire-and-forget: any failure (e.g. bundled model missing) just means the
 * first real call pays the cost/surfaces the error as it always did.
 */
export function warmEmbeddings(): void {
  void embedTexts(['warm-up']).catch(() => {})
}

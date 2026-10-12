import * as path from 'path'
import { Worker } from 'worker_threads'
import type { BidDocumentInput } from '../core/bidDocument'

export interface BidDocumentBatchEntry {
  templatePath: string
  kind: 'ee' | 'se'
  input: BidDocumentInput
  filePath: string
}

type WorkerMsg =
  | { type: 'progress'; done: number; total: number }
  | { type: 'done'; written: string[] }
  | { type: 'error'; error: string }

/**
 * Fills/sanitizes/writes a whole batch of Bid Documents on a worker thread
 * (electron/bidDocumentWorker.ts), not the main process — see that file's
 * comment. One worker per batch; it exits once the batch settles.
 */
export function runBidDocumentBatch(
  entries: BidDocumentBatchEntry[],
  onProgress: (done: number, total: number) => void
): Promise<string[]> {
  return new Promise<string[]>((resolve, reject) => {
    const worker = new Worker(path.join(__dirname, 'bidDocumentWorker.js'))
    worker.on('message', (msg: WorkerMsg) => {
      if (msg.type === 'progress') onProgress(msg.done, msg.total)
      else if (msg.type === 'done') {
        resolve(msg.written)
        void worker.terminate()
      } else if (msg.type === 'error') {
        reject(new Error(msg.error))
        void worker.terminate()
      }
    })
    worker.on('error', (err) => reject(err))
    worker.postMessage({ entries })
  })
}

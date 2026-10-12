import { parentPort } from 'worker_threads'
import * as fs from 'fs'
import { fillBidDocument, fillSeBidDocument } from '../core/bidDocument'
import { sanitizeDocxForWord2007 } from '../core/word2007Compat'
import type { BidDocumentInput } from '../core/bidDocument'

// Runs on its own OS thread — filling+sanitizing a docx is CPU-bound and
// synchronous, so doing a whole batch of them on the main process's single
// thread (even with a setImmediate yield between documents) still hogs it
// for the total duration, making every window feel unresponsive. The main
// thread resolves each entry's template path + Electron save-dialog/
// filename work (cheap) and hands this worker the actual fill/sanitize/write.

interface Entry {
  templatePath: string
  kind: 'ee' | 'se'
  input: BidDocumentInput
  filePath: string
}

interface Request {
  entries: Entry[]
}

parentPort?.on('message', (msg: Request) => {
  try {
    const templateCache = new Map<string, Buffer>()
    const written: string[] = []
    for (let i = 0; i < msg.entries.length; i++) {
      const entry = msg.entries[i]
      let templateBuffer = templateCache.get(entry.templatePath)
      if (!templateBuffer) {
        templateBuffer = fs.readFileSync(entry.templatePath)
        templateCache.set(entry.templatePath, templateBuffer)
      }
      const fill = entry.kind === 'se' ? fillSeBidDocument : fillBidDocument
      const filled = sanitizeDocxForWord2007(fill(templateBuffer, entry.input))
      fs.writeFileSync(entry.filePath, filled)
      written.push(entry.filePath)
      parentPort?.postMessage({ type: 'progress', done: i + 1, total: msg.entries.length })
    }
    parentPort?.postMessage({ type: 'done', written })
  } catch (e) {
    parentPort?.postMessage({ type: 'error', error: e instanceof Error ? e.message : String(e) })
  }
})

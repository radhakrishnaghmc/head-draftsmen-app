import * as path from 'path'
import { fork } from 'child_process'
import type { SplitProgress } from './excelSplit'

type WorkerMsg =
  | { type: 'progress'; done: number; total: number; sheet: string }
  | { type: 'done'; files: string[] }
  | { type: 'error'; error: string }

/**
 * Split a workbook in an isolated child process (electron/splitWorker.ts) with a
 * raised V8 heap, so an out-of-memory on a very large workbook only kills the
 * child — surfaced here as a rejected promise — rather than hard-crashing the
 * whole app (as it would if the split ran in the main process). ELECTRON_RUN_AS_NODE
 * launches the Electron binary as plain Node so the child works in a packaged app.
 */
export function runSplitInWorker(
  srcPath: string,
  outDir: string,
  sheetNames: string[] | null,
  onProgress: SplitProgress
): Promise<string[]> {
  return new Promise<string[]>((resolve, reject) => {
    const child = fork(path.join(__dirname, 'splitWorker.js'), [], {
      env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' },
      execArgv: ['--max-old-space-size=8192']
    })
    let settled = false
    const finish = (fn: () => void): void => {
      if (settled) return
      settled = true
      fn()
      child.kill()
    }

    child.on('message', (msg: WorkerMsg) => {
      if (msg.type === 'progress') onProgress(msg.done, msg.total, msg.sheet)
      else if (msg.type === 'done') finish(() => resolve(msg.files))
      else if (msg.type === 'error') finish(() => reject(new Error(msg.error)))
    })
    // Exit without a done/error message means the child died — almost always the
    // workbook being too large to split within even the raised heap.
    child.on('exit', (code, signal) =>
      finish(() =>
        reject(
          new Error(
            `The workbook was too large to separate — the split ran out of memory (exit ${code ?? signal}). ` +
              `Try separating one sheet at a time, or split it in Excel.`
          )
        )
      )
    )
    child.on('error', (err) => finish(() => reject(err)))
    child.send({ srcPath, outDir, sheetNames })
  })
}

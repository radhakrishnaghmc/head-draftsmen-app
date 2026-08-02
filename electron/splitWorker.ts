// Runs the (memory-heavy) workbook split in an isolated child process — see
// electron/splitRunner.ts. Splitting a large multi-sheet workbook via ExcelJS
// can exhaust the heap; in a separate process a "JavaScript heap out of memory"
// only kills this child (the parent reports it as an error), instead of hard-
// crashing the whole app the way it would in the main process. The child is
// forked with a raised --max-old-space-size so most real workbooks fit.
import { splitWorkbookSheets } from './excelSplit'

interface SplitJob {
  srcPath: string
  outDir: string
  sheetNames: string[] | null
}

process.on('message', (msg: SplitJob) => {
  void (async () => {
    try {
      const files = await splitWorkbookSheets(
        msg.srcPath,
        msg.outDir,
        (done, total, sheet) => process.send?.({ type: 'progress', done, total, sheet }),
        msg.sheetNames ?? undefined
      )
      process.send?.({ type: 'done', files })
    } catch (e) {
      process.send?.({ type: 'error', error: e instanceof Error ? e.message : String(e) })
    }
  })()
})

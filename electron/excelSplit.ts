import * as ExcelJS from 'exceljs'
import * as fs from 'fs'
import * as path from 'path'
import { stripDataValidations } from '../core/templateWorkbook'

export interface SplitResult {
  dir: string
  files: string[]
}

/** Called as each sheet is written, so a caller can drive a progress bar. */
export type SplitProgress = (done: number, total: number, sheet: string) => void

// Turn a sheet name into a safe file name (Excel tab names allow characters a
// file name can't). Falls back to "Sheet" if nothing usable remains.
function sanitizeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim() || 'Sheet'
}

// Hand control back to the event loop between sheets so the main process stays
// responsive to other IPC while a large workbook is being split.
const yieldToLoop = (): Promise<void> => new Promise((resolve) => setImmediate(resolve))

/**
 * Split every sheet of one workbook into its own single-sheet .xlsx, written
 * into `outDir` and named after the sheet's tab. Each output keeps that
 * sheet's values, formulas, merges and formatting.
 *
 * The workbook is parsed **once** and each output is built by copying a single
 * sheet's model into a fresh one-sheet workbook. (The previous implementation
 * re-read the whole file from disk for every sheet — an O(sheets × parse) cost
 * that made large monitoring workbooks, which carry data-validation dropdowns
 * on every sheet, appear to hang.) Workbook media is shared across the outputs
 * so any anchored images still resolve. stripDataValidations is applied for the
 * same reason the template code applies it (ExcelJS mis-writes large validation
 * ranges into an Excel "repair" prompt). Colliding tab names get a " (2)"
 * suffix. Returns the saved file paths in sheet order.
 */
export async function splitWorkbookSheets(
  srcPath: string,
  outDir: string,
  onProgress?: SplitProgress
): Promise<string[]> {
  const src = new ExcelJS.Workbook()
  await src.xlsx.readFile(srcPath)
  const sheets = [...src.worksheets]
  const total = sheets.length
  // Images live at workbook level and worksheet cells reference them by index;
  // reuse the same media array in every output so those references stay valid.
  const media = (src as unknown as { media: unknown[] }).media

  const used = new Set<string>()
  const written: string[] = []

  for (let i = 0; i < sheets.length; i++) {
    const ws = sheets[i]
    onProgress?.(i, total, ws.name)

    const out = new ExcelJS.Workbook()
    out.creator = src.creator
    out.created = src.created
    if (media) (out as unknown as { media: unknown[] }).media = media
    const newWs = out.addWorksheet(ws.name)
    // Carry over the full sheet content (cells, styles, merges, column/row
    // sizing, page setup, images) via the worksheet model, keeping its name.
    // The model *getter* emits merged ranges as `merges`, but the *setter*
    // restores them from `mergeCells` — so map one to the other or merges are
    // silently dropped on copy.
    const model = ws.model as typeof ws.model & { merges?: string[] }
    newWs.model = { ...model, name: ws.name, mergeCells: model.merges ?? [] } as typeof ws.model
    stripDataValidations(newWs)

    const base = sanitizeFileName(ws.name)
    let fileName = `${base}.xlsx`
    let n = 2
    while (used.has(fileName.toLowerCase()) || fs.existsSync(path.join(outDir, fileName))) {
      fileName = `${base} (${n}).xlsx`
      n += 1
    }
    used.add(fileName.toLowerCase())

    const outPath = path.join(outDir, fileName)
    await out.xlsx.writeFile(outPath)
    written.push(outPath)

    onProgress?.(i + 1, total, ws.name)
    await yieldToLoop()
  }
  return written
}

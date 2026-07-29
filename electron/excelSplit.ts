import * as ExcelJS from 'exceljs'
import * as fs from 'fs'
import * as path from 'path'
import { stripDataValidations } from '../core/templateWorkbook'

export interface SplitResult {
  dir: string
  files: string[]
}

// Turn a sheet name into a safe file name (Excel tab names allow characters a
// file name can't). Falls back to "Sheet" if nothing usable remains.
function sanitizeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, ' ').replace(/\s+/g, ' ').trim() || 'Sheet'
}

/**
 * Split every sheet of one workbook into its own single-sheet .xlsx, written
 * into `outDir` and named after the sheet's tab. Each output keeps that
 * sheet's values, formulas, merges and formatting — the workbook is re-loaded
 * per sheet and every *other* sheet removed, so nothing about the kept sheet
 * is regenerated beyond ExcelJS's normal round-trip. stripDataValidations is
 * applied for the same reason the template code applies it (ExcelJS mis-writes
 * large validation ranges into an Excel "repair" prompt). Colliding tab names
 * get a " (2)" suffix. Returns the saved file paths in sheet order.
 */
export async function splitWorkbookSheets(srcPath: string, outDir: string): Promise<string[]> {
  const probe = new ExcelJS.Workbook()
  await probe.xlsx.readFile(srcPath)
  const names = probe.worksheets.map((w) => w.name)

  const used = new Set<string>()
  const written: string[] = []
  for (const name of names) {
    const wb = new ExcelJS.Workbook()
    await wb.xlsx.readFile(srcPath)
    for (const ws of [...wb.worksheets]) {
      if (ws.name === name) stripDataValidations(ws)
      else wb.removeWorksheet(ws.id)
    }

    const base = sanitizeFileName(name)
    let fileName = `${base}.xlsx`
    let n = 2
    while (used.has(fileName.toLowerCase()) || fs.existsSync(path.join(outDir, fileName))) {
      fileName = `${base} (${n}).xlsx`
      n += 1
    }
    used.add(fileName.toLowerCase())

    const outPath = path.join(outDir, fileName)
    await wb.xlsx.writeFile(outPath)
    written.push(outPath)
  }
  return written
}

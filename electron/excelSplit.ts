import * as ExcelJS from 'exceljs'
import * as fs from 'fs'
import * as path from 'path'
import PizZip from 'pizzip'
import { stripDataValidations } from '../core/templateWorkbook'

/**
 * Strip the memory-bomb parts out of a workbook before ExcelJS parses it, and
 * return the slimmed .xlsx as a buffer. Some workbooks accumulate tens of
 * thousands of junk `<definedName>`s and hundreds of external-workbook links
 * (from cross-workbook copy/paste); ExcelJS materialises all of them into
 * objects, which can balloon to many GB and OOM on a file only a few MB on
 * disk. None of it is needed to separate sheets — cached cell values are kept —
 * so we drop the external links, the defined names, the external references and
 * the calc chain at the zip level (cheap byte edits) first.
 */
function slimWorkbookBuffer(srcPath: string): Buffer {
  const zip = new PizZip(fs.readFileSync(srcPath))
  for (const name of Object.keys(zip.files)) {
    if (/^xl\/externalLinks\//.test(name)) zip.remove(name)
  }
  const wbFile = zip.file('xl/workbook.xml')
  if (wbFile) {
    const wb = wbFile
      .asText()
      .replace(/<definedNames>[\s\S]*?<\/definedNames>/g, '')
      .replace(/<externalReferences>[\s\S]*?<\/externalReferences>/g, '')
    zip.file('xl/workbook.xml', wb)
  }
  const relsFile = zip.file('xl/_rels/workbook.xml.rels')
  if (relsFile) {
    zip.file(
      'xl/_rels/workbook.xml.rels',
      relsFile.asText().replace(/<Relationship\b[^>]*externalLink[^>]*\/>/g, '')
    )
  }
  if (zip.file('xl/calcChain.xml')) zip.remove('xl/calcChain.xml')
  const ctFile = zip.file('[Content_Types].xml')
  if (ctFile) {
    zip.file(
      '[Content_Types].xml',
      ctFile
        .asText()
        .replace(/<Override\b[^>]*externalLink[^>]*\/>/g, '')
        .replace(/<Override\b[^>]*calcChain[^>]*\/>/g, '')
    )
  }
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&amp;/g, '&')
}

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
/**
 * The worksheet names in a workbook, in tab order — for letting the user pick
 * which to separate. Reads only the tiny `xl/workbook.xml` out of the .xlsx zip
 * (the sheet list lives there), never parsing the whole workbook — a full
 * ExcelJS parse of a large multi-sheet file just to read its tab names can
 * exhaust memory and crash the process.
 */
export function readWorkbookSheetNames(srcPath: string): string[] {
  const zip = new PizZip(fs.readFileSync(srcPath))
  const xml = zip.file('xl/workbook.xml')?.asText()
  if (!xml) return []
  const names: string[] = []
  const re = /<(?:x:)?sheet\b[^>]*\bname="([^"]*)"/g
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) names.push(decodeXmlEntities(m[1]))
  return names
}

/**
 * Replace every formula cell with its cached result value ("paste values only").
 * A separated sheet is standalone, so any formula that referenced another sheet
 * or an external workbook would otherwise resolve to nothing / #REF and the cell
 * would look empty — losing the data. Flattening to the value ExcelJS already
 * read keeps the numbers the sheet was showing. (When a formula has no cached
 * result the cell is cleared rather than left as a broken formula.)
 */
function flattenFormulasToValues(ws: ExcelJS.Worksheet): void {
  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value as unknown
      if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
        cell.value = (v as { result?: ExcelJS.CellValue }).result ?? null
      }
    })
  })
}

export async function splitWorkbookSheets(
  srcPath: string,
  outDir: string,
  onProgress?: SplitProgress,
  // When given (and non-empty), only these sheets are written; otherwise all are.
  sheetNames?: string[]
): Promise<string[]> {
  const src = new ExcelJS.Workbook()
  // Load from a slimmed buffer (external links / defined names removed) rather
  // than the raw file, so a workbook carrying a huge junk-reference payload
  // doesn't OOM ExcelJS. See slimWorkbookBuffer.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any -- @types/node's
  // generic Buffer doesn't line up with ExcelJS's older Buffer param type.
  await src.xlsx.load(slimWorkbookBuffer(srcPath) as any)
  const wanted = sheetNames && sheetNames.length > 0 ? new Set(sheetNames) : null
  const sheets = wanted ? src.worksheets.filter((ws) => wanted.has(ws.name)) : [...src.worksheets]
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
    flattenFormulasToValues(newWs)

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

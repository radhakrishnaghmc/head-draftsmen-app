import * as ExcelJS from 'exceljs'
import PizZip from 'pizzip'
import type { CellEdit } from './technicalSanction'
import { stripDataValidations } from './templateWorkbook'

const GREEN_ARGB = 'FF008000'
const RED_ARGB = 'FFCC0000'

/**
 * Real-world estimate workbooks accumulate "phantom" bloat over years of
 * copy-pasting cells from other files — one seen in testing carried 520
 * external-link parts (23MB of embedded XML) *and* 32,718 defined names
 * (2.3MB of workbook.xml, mostly auto-generated garbage names from repeated
 * paste operations), despite the actual sheet holding barely a thousand
 * cells. ExcelJS builds a full in-memory model of every part it loads, and
 * parsing that much external-link/defined-name data blows past several GB
 * of heap and crashes before the workbook even finishes loading — nothing to
 * do with the size of the visible sheet. Since editing specific cells and
 * adding a rate-analysis sheet never needs to resolve either of these,
 * strip them from the zip (and the dangling references to them in
 * workbook.xml, its rels, and [Content_Types].xml) before ExcelJS ever sees
 * the file.
 */
export function stripWorkbookBloat(buffer: Buffer): Buffer {
  const zip = new PizZip(buffer)
  const hasExternalLinks = zip.file(/^xl\/externalLinks\//).length > 0

  const workbookPath = 'xl/workbook.xml'
  const workbookXml = zip.file(workbookPath)?.asText()
  const hasDefinedNames = !!workbookXml && /<definedNames>/.test(workbookXml)
  if (!hasExternalLinks && !hasDefinedNames) return buffer

  if (hasExternalLinks) {
    zip.remove('xl/externalLinks')

    const relsPath = 'xl/_rels/workbook.xml.rels'
    const relsXml = zip.file(relsPath)?.asText()
    if (relsXml) {
      zip.file(relsPath, relsXml.replace(/<Relationship[^>]*Type="[^"]*\/externalLink"[^>]*\/>/g, ''))
    }

    const contentTypesPath = '[Content_Types].xml'
    const contentTypesXml = zip.file(contentTypesPath)?.asText()
    if (contentTypesXml) {
      zip.file(
        contentTypesPath,
        contentTypesXml.replace(/<Override[^>]*PartName="\/xl\/externalLinks\/[^"]*"[^>]*\/>/g, '')
      )
    }
  }

  if (workbookXml) {
    let cleaned = workbookXml
    if (hasExternalLinks) cleaned = cleaned.replace(/<externalReferences>.*?<\/externalReferences>/, '')
    if (hasDefinedNames) cleaned = cleaned.replace(/<definedNames>.*?<\/definedNames>/, '')
    zip.file(workbookPath, cleaned)
  }

  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

// Reassign the whole `cell.style` object rather than mutating `.font`
// directly — ExcelJS shares one style object by reference across every cell
// that used the same style index in the loaded file, so mutating a
// sub-property in place silently recolors every other cell sharing that
// style too. See core/boqTemplate.ts's styleCell() for the same fix.
function setCellColor(cell: ExcelJS.Cell, color: 'green' | 'red'): void {
  cell.style = {
    ...cell.style,
    font: { ...cell.style.font, color: { argb: color === 'green' ? GREEN_ARGB : RED_ARGB } }
  }
}

/**
 * Apply Technical Sanction cell edits to the uploaded estimate file itself
 * (not a bundled template) — updating rate values and coloring description
 * cells green/red — while preserving everything else about the original
 * file's formatting, formulas, and layout. When `rateAnalysisRows` is
 * non-empty, it's written to the workbook's second sheet (appended if one
 * already exists, e.g. a pre-existing manual rate-analysis sheet; created as
 * "Sheet 2" otherwise) — the supporting material/labour/machinery buildup
 * behind each rate actually applied, matching the convention seen in real
 * sanctioned estimates.
 */
export async function applyTechnicalSanctionEdits(
  originalBuffer: Buffer,
  sheetName: string,
  edits: CellEdit[],
  rateAnalysisRows: (string | number)[][] = []
): Promise<Buffer> {
  const cleanedBuffer = stripWorkbookBloat(originalBuffer)
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(cleanedBuffer as unknown as ArrayBuffer)
  const ws = workbook.getWorksheet(sheetName) ?? workbook.worksheets[0]
  if (!ws) throw new Error('Could not find a worksheet to edit in the estimate file.')
  stripDataValidations(ws)

  for (const edit of edits) {
    // ExcelJS rows/columns are 1-indexed; our grid positions are 0-indexed.
    const cell = ws.getCell(edit.row + 1, edit.col + 1)
    if (edit.value !== undefined) cell.value = edit.value
    if (edit.color) setCellColor(cell, edit.color)
  }

  if (rateAnalysisRows.length > 0) {
    const analysisSheet = workbook.worksheets[1] ?? workbook.addWorksheet('Sheet 2')
    for (const row of rateAnalysisRows) analysisSheet.addRow(row)
  }

  const out = await workbook.xlsx.writeBuffer()
  return Buffer.from(out)
}

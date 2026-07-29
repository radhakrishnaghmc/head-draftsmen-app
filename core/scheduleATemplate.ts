import * as ExcelJS from 'exceljs'
import type { ScheduleAItem, ScheduleAMeta } from './scheduleA'
import { indianDigitGroups } from './worksAmounts'
import { stripDataValidations } from './templateWorkbook'

// Fixed column layout of the bundled Schedule-A template's item table.
const ITEM_NO_COL = 1
const QTY_COL = 2
const DESC_COL = 3
const RATE_COL = 4
const UNITS_COL = 5
const AMOUNT_COL = 6

function cellText(v: ExcelJS.CellValue): string {
  if (v == null) return ''
  if (typeof v === 'object') {
    const rich = v as { richText?: { text: string }[] }
    if (rich.richText) return rich.richText.map((r) => r.text).join('')
    const formulaVal = v as { result?: unknown }
    if ('result' in (v as object)) return String(formulaVal.result ?? '')
  }
  return String(v)
}

function numOrNull(s: string | undefined): number | null {
  if (!s) return null
  const n = Number(String(s).replace(/[,%]/g, '').trim())
  return Number.isFinite(n) && String(s).trim() !== '' ? n : null
}

/**
 * The bundled Schedule-A template is a real filled-in sample, so it carries a
 * specific office in two places the metadata fill doesn't touch: the signature
 * block's circle line ("Gajularamaram Circle-57,CMC") and the preamble's zone
 * ("…the Quthbullapur Circle does not accept…"). Swap those sample values for
 * the current work's circle/zone so a saved Schedule A isn't stamped with the
 * wrong office. Applied to plain-string and rich-text cells alike; the zone is
 * only substituted when it's a real name (never a bare number).
 */
function replaceOfficeText(ws: ExcelJS.Worksheet, meta?: ScheduleAMeta): void {
  const circle = meta?.circle?.trim()
  const zone = meta?.zone?.trim()
  const zoneOk = !!zone && !/^\d+$/.test(zone)
  if (!circle && !zoneOk) return

  const swap = (text: string): string => {
    let s = text
    if (circle) s = s.replace(/Gajularamaram\s+Circle\s*-?\s*57/gi, circle)
    if (zoneOk) s = s.replace(/Quthbullapur(?=\s+Circle)/gi, zone!)
    return s
  }

  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value
      if (typeof v === 'string') {
        const s = swap(v)
        if (s !== v) cell.value = s
      } else if (v && typeof v === 'object' && 'richText' in v && Array.isArray((v as { richText?: unknown }).richText)) {
        const rich = v as { richText: { text: string }[] }
        let changed = false
        const runs = rich.richText.map((run) => {
          const t = swap(run.text)
          if (t !== run.text) changed = true
          return { ...run, text: t }
        })
        if (changed) cell.value = { richText: runs } as ExcelJS.CellValue
      }
    })
  })
}

/**
 * Fill the bundled Schedule-A template in place — preserving its fonts,
 * borders, merged cells and formulas — instead of rebuilding the document
 * from scratch. Item rows are inserted or removed to match the uploaded
 * BOQ's row count, using the template's own item row as the style/formula
 * reference, and its Amount cells stay live formulas (qty × rate, and a
 * SUM total) rather than static numbers, matching the original design.
 */
export async function fillScheduleATemplate(
  templateBuffer: Buffer,
  items: ScheduleAItem[],
  meta?: ScheduleAMeta
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(templateBuffer as unknown as ArrayBuffer)
  const ws = workbook.worksheets[0]
  if (!ws) throw new Error('Schedule A template has no sheet.')
  stripDataValidations(ws)

  let headerRow = -1
  for (let r = 1; r <= Math.min(ws.rowCount, 40); r++) {
    if (/^item\s*no\.?$/i.test(cellText(ws.getCell(r, ITEM_NO_COL).value).trim())) {
      headerRow = r
      break
    }
  }
  if (headerRow === -1) {
    throw new Error('Could not find the item table header row ("Item No.") in the Schedule A template.')
  }

  const firstItemRow = headerRow + 2 // header row, then the 1..6 numbering row
  let templateItemCount = 0
  while (cellText(ws.getCell(firstItemRow + templateItemCount, DESC_COL).value).trim() !== '') {
    templateItemCount += 1
  }
  if (templateItemCount === 0) {
    throw new Error('Schedule A template has no item row to use as a formatting reference.')
  }

  const n = items.length
  if (n > templateItemCount) {
    ws.duplicateRow(firstItemRow + templateItemCount - 1, n - templateItemCount, true)
  } else if (n < templateItemCount) {
    ws.spliceRows(firstItemRow + n, templateItemCount - n)
  }

  for (let i = 0; i < n; i++) {
    const r = firstItemRow + i
    const it = items[i]
    ws.getCell(r, ITEM_NO_COL).value = i + 1
    ws.getCell(r, QTY_COL).value = numOrNull(it.quantity) ?? it.quantity
    ws.getCell(r, DESC_COL).value = it.description
    ws.getCell(r, RATE_COL).value = numOrNull(it.rate) ?? it.rate
    ws.getCell(r, UNITS_COL).value = it.units
    ws.getCell(r, AMOUNT_COL).value = { formula: `ROUND(B${r}*D${r},0)` }
  }

  const totalRow = firstItemRow + n
  ws.getCell(totalRow, AMOUNT_COL).value = n > 0 ? { formula: `SUM(F${firstItemRow}:F${totalRow - 1})` } : 0

  const total = items.reduce((sum, it) => sum + (Number(String(it.amount).replace(/,/g, '')) || 0), 0)
  // Fallback only when there's no Works List match at all — the item table's
  // own total, Indian-formatted the same way meta.estimateAmount/ecvAmount are
  // (metaFromWorksRow in core/scheduleA.ts already applies the app's
  // Lakhs-to-rupees + Indian-numbering money rules, so those come in
  // pre-formatted, e.g. "45,00,000/-" — not run through numOrNull). Never
  // used to paper over a matched row's genuinely blank ECV — that stays
  // blank rather than being shown as if it were a real figure.
  const totalDisplay = meta ? '' : `${indianDigitGroups(total)}/-`

  // Metadata rows sit above the item header — match by label text so this
  // keeps working even if a row or two shifts in a future template edit.
  for (let r = 1; r < headerRow; r++) {
    const label = cellText(ws.getCell(r, 1).value).trim().toLowerCase()
    if (label.startsWith('name of work')) {
      ws.getCell(r, 1).value = meta?.nameOfWork ? `Name of work: ${meta.nameOfWork}` : 'Name of work:'
    } else if (label.startsWith('name of the contractor')) {
      ws.getCell(r, 1).value = meta?.contractorName
        ? `Name of the Contractor : ${meta.contractorName}`
        : 'Name of the Contractor : '
    } else if (label.startsWith('estimate amount')) {
      ws.getCell(r, 3).value = meta?.estimateAmount ?? totalDisplay
    } else if (label.startsWith('ecv amount')) {
      ws.getCell(r, 3).value = meta?.ecvAmount ?? totalDisplay
    } else if (label.startsWith('contract amount')) {
      ws.getCell(r, 3).value = meta?.contractAmount ?? ''
    }
  }

  // Tender-quoted % and its "Less:" restatement — filled from the Works List
  // when we have it (meta.tenderPercentage), otherwise left blank.
  const tenderPct = numOrNull(meta?.tenderPercentage)
  let tenderQuotedRow = -1
  let inFiguresRow = -1
  for (let r = totalRow; r <= ws.rowCount; r++) {
    const c1 = cellText(ws.getCell(r, 1).value).trim()
    const c3 = cellText(ws.getCell(r, 3).value).trim()
    if (/^tender quoted$/i.test(c3)) {
      tenderQuotedRow = r
      ws.getCell(r, 4).value = tenderPct
    } else if (/^less:/i.test(c3)) {
      ws.getCell(r, 3).value =
        tenderPct != null
          ? `Less:                       ( -${tenderPct} )% Less`
          : 'Less:                       (        )% Less'
    } else if (/in figures/i.test(c1)) {
      inFiguresRow = r
    }
  }
  // Row insertion/removal above shifted the total and tender-quoted rows —
  // the "In Figures" formula must point at their new positions, not the
  // template's original (now stale) row numbers.
  if (inFiguresRow !== -1 && tenderQuotedRow !== -1) {
    ws.getCell(inFiguresRow, 4).value = {
      formula: `+F${totalRow}-(F${totalRow}*D${tenderQuotedRow}%)`
    }
  }

  replaceOfficeText(ws, meta)

  const out = await workbook.xlsx.writeBuffer()
  return Buffer.from(out)
}

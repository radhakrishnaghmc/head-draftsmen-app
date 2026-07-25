import * as ExcelJS from 'exceljs'
import type { MaterialTotals } from './materialEstimate'
import { stripDataValidations } from './templateWorkbook'

// Fixed layout of the bundled Material Estimation Template
// (resources/material-estimation-template.xlsx) — a small, fixed-row form,
// not an item table that grows with the estimate: header fields at rows
// 3-6, and exactly 7 material rows (9-15) in a fixed order.
const WORK_NAME_ROW = 3
const DEPARTMENT_ROW = 4
const DISTRICT_ROW = 5
const ECV_ROW = 6
const HEADER_VALUE_COL = 3 // C3:E3 etc, merged

// Row order matches the template's own S.No 1-7 exactly.
const STONE_AGGREGATES_ROW = 9
const SAND_ROW = 10
const GRAVEL_ROW = 11
const GRANITE_ROW = 12
const NAPA_ROW = 13
const CEMENT_ROW = 14
const STEEL_ROW = 15
const QTY_COL = 5 // "Total Quantity Required"

export interface MaterialEstimateMeta {
  workName?: string
  departmentName?: string
  district?: string
  /** In rupees — matches how the template's own "Estimated Contract Value (ECV)" field is worded. */
  ecvRupees?: number
}

/** Rounded to 2 decimals — a raw derived quantity (e.g. 3.14159265 MT) reads as a spurious level of precision on a document meant to be checked against the department's own figures. */
function round2(n: number): number {
  return Math.round(n * 100) / 100
}

export async function fillMaterialTemplate(
  templateBuffer: Buffer,
  totals: MaterialTotals,
  meta: MaterialEstimateMeta = {}
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(templateBuffer as unknown as ArrayBuffer)
  const ws = workbook.worksheets[0]
  if (!ws) throw new Error('Material Estimation template has no sheet.')
  stripDataValidations(ws)

  if (meta.workName) ws.getCell(WORK_NAME_ROW, HEADER_VALUE_COL).value = meta.workName
  if (meta.departmentName) ws.getCell(DEPARTMENT_ROW, HEADER_VALUE_COL).value = meta.departmentName
  if (meta.district) ws.getCell(DISTRICT_ROW, HEADER_VALUE_COL).value = meta.district
  if (meta.ecvRupees != null) ws.getCell(ECV_ROW, HEADER_VALUE_COL).value = meta.ecvRupees

  ws.getCell(STONE_AGGREGATES_ROW, QTY_COL).value = round2(totals.stoneAggregatesMt)
  ws.getCell(SAND_ROW, QTY_COL).value = round2(totals.sandMt)
  ws.getCell(GRAVEL_ROW, QTY_COL).value = round2(totals.gravelMt)
  ws.getCell(GRANITE_ROW, QTY_COL).value = round2(totals.graniteSqft)
  ws.getCell(NAPA_ROW, QTY_COL).value = round2(totals.napaSqft)
  ws.getCell(CEMENT_ROW, QTY_COL).value = round2(totals.cementMt)
  ws.getCell(STEEL_ROW, QTY_COL).value = round2(totals.steelMt)

  const out = await workbook.xlsx.writeBuffer()
  return Buffer.from(out)
}

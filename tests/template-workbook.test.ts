import { describe, expect, it } from 'vitest'
import * as ExcelJS from 'exceljs'
import { stripDataValidations, trimToContent } from '../core/templateWorkbook'

// ExcelJS's reader expands a validation's sqref into one model entry *per
// cell* it covers; for a rule spanning down to row 1048576 that model can
// balloon to millions of entries, and its writer's own re-coalescing logic
// is buggy at that scale — it emits overlapping/duplicated ranges for what
// was originally one clean rule, which real Excel refuses to open without a
// repair prompt. See core/templateWorkbook.ts's own doc comment.
async function roundTripDataValidationCount(): Promise<{ before: number; after: number }> {
  const workbook = new ExcelJS.Workbook()
  const ws = workbook.addWorksheet('Sheet1')
  ws.getCell('A1').value = 'header'
  // A rule spanning a large row range, same shape as the real bundled
  // templates that triggered this bug (resources/boq-template.xlsx,
  // resources/deviation-template.xlsx both have rules like this).
  ;(ws as unknown as { dataValidations: { add(address: string, rule: object): void } }).dataValidations.add(
    'A1:A100000',
    { type: 'decimal', operator: 'greaterThan', formulae: [0] }
  )

  const before = await workbook.xlsx.writeBuffer()
  const beforeWb = new ExcelJS.Workbook()
  await beforeWb.xlsx.load(before as unknown as ArrayBuffer)
  const beforeCount = Object.keys(
    (beforeWb.worksheets[0] as unknown as { dataValidations: { model: object } }).dataValidations.model
  ).length

  stripDataValidations(workbook.worksheets[0])
  const after = await workbook.xlsx.writeBuffer()
  const afterWb = new ExcelJS.Workbook()
  await afterWb.xlsx.load(after as unknown as ArrayBuffer)
  const afterCount = Object.keys(
    (afterWb.worksheets[0] as unknown as { dataValidations: { model: object } }).dataValidations.model
  ).length

  return { before: beforeCount, after: afterCount }
}

describe('stripDataValidations', () => {
  it('removes every data-validation rule so a large-range rule can never trigger the round-trip corruption bug', async () => {
    const { before, after } = await roundTripDataValidationCount()
    expect(before).toBeGreaterThan(1) // confirms the large-range model expansion is real, not a no-op in this test setup
    expect(after).toBe(0)
  })
})

describe('trimToContent', () => {
  it('truncates rows all the way to the worksheet\'s own last row — the exact case worksheet.spliceRows() silently no-ops on', async () => {
    const workbook = new ExcelJS.Workbook()
    const ws = workbook.addWorksheet('Sheet1')
    // A cell far out, styled but with no value — the same "phantom" shape
    // found in the real bundled templates (selecting/formatting a whole
    // row or column at once in Excel leaves styled-but-empty cells behind).
    ws.getCell(1, 1).value = 'header'
    ws.getCell(50, 200).value = 'phantom'
    expect(ws.rowCount).toBe(50)
    expect(ws.columnCount).toBe(200)

    trimToContent(ws, 1, 1)

    expect(ws.rowCount).toBe(1)
    const out = await workbook.xlsx.writeBuffer()
    const reloaded = new ExcelJS.Workbook()
    await reloaded.xlsx.load(out as unknown as ArrayBuffer)
    const reloadedWs = reloaded.worksheets[0]
    expect(reloadedWs.rowCount).toBe(1)
    // The phantom cell's own value/style must be gone from the written
    // file too, not just hidden behind a stale rowCount/columnCount getter.
    expect(reloadedWs.getCell(50, 200).value).toBeNull()
  })
})

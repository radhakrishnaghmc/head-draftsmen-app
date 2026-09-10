import { describe, it, expect, afterAll } from 'vitest'
import * as ExcelJS from 'exceljs'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { readExcelGrid, allSheetGridsFromBuffer } from '../core/excel'

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'hidden-rows-test-'))

afterAll(() => fs.rmSync(tmp, { recursive: true, force: true }))

describe('gridFromSheet (via readExcelGrid / allSheetGridsFromBuffer)', () => {
  it('drops a manually-hidden row from the grid entirely — a BOQ/estimate built from it must not count that item', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('Estimate')
    ws.getCell('A1').value = 'S.No'
    ws.getCell('B1').value = 'Description'
    ws.getCell('A2').value = '1'
    ws.getCell('B2').value = 'Earthwork'
    ws.getCell('A3').value = '2'
    ws.getCell('B3').value = 'Superseded item — hidden by the estimator'
    ws.getRow(3).hidden = true
    ws.getCell('A4').value = '3'
    ws.getCell('B4').value = 'Concreting'

    const file = path.join(tmp, 'hidden.xlsx')
    await wb.xlsx.writeFile(file)

    const { grid } = readExcelGrid(file)
    const descriptions = grid.map((row) => row[1])
    expect(descriptions).toEqual(['Description', 'Earthwork', 'Concreting'])
    expect(descriptions).not.toContain('Superseded item — hidden by the estimator')
  })

  it('keeps a merely-blank row (not hidden) — only hidden rows are excluded', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('Estimate')
    ws.getCell('A1').value = 'S.No'
    ws.getCell('A2').value = '1'
    // Row 3 left entirely blank, not hidden.
    ws.getCell('A4').value = '2'

    const file = path.join(tmp, 'blank-row.xlsx')
    await wb.xlsx.writeFile(file)

    const { grid } = readExcelGrid(file)
    expect(grid.map((row) => row[0])).toEqual(['S.No', '1', '', '2'])
  })

  it('applies the same exclusion when reading every sheet (allSheetGridsFromBuffer, the multi-sheet estimate flow)', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('Sheet1')
    ws.getCell('A1').value = 'kept'
    ws.getCell('A2').value = 'hidden'
    ws.getRow(2).hidden = true
    ws.getCell('A3').value = 'also kept'

    const buf = Buffer.from(await wb.xlsx.writeBuffer())
    const grids = allSheetGridsFromBuffer(buf, 'wb.xlsx', '')
    expect(grids[0].grid.map((row) => row[0])).toEqual(['kept', 'also kept'])
  })
})

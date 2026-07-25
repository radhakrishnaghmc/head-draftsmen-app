import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import * as ExcelJS from 'exceljs'
import { fillMaterialTemplate } from '../core/materialTemplate'
import type { MaterialTotals } from '../core/materialEstimate'

const TEMPLATE_PATH = resolve(__dirname, '../resources/material-estimation-template.xlsx')

const totals: MaterialTotals = {
  stoneAggregatesMt: 12.345,
  sandMt: 6.789,
  gravelMt: 3.5,
  graniteSqft: 1000.001,
  napaSqft: 250.5,
  cementMt: 4.444,
  steelMt: 1.111
}

describe('fillMaterialTemplate', () => {
  it('fills the header meta and the 7 material rows in the template\'s own fixed order, rounded to 2 decimals', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = await fillMaterialTemplate(buffer, totals, {
      workName: 'Laying of CC road',
      departmentName: 'Municipal Administration',
      district: 'Medchal-Malkajgiri',
      ecvRupees: 1234567
    })

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]

    expect(ws.getCell(3, 3).value).toBe('Laying of CC road')
    expect(ws.getCell(4, 3).value).toBe('Municipal Administration')
    expect(ws.getCell(5, 3).value).toBe('Medchal-Malkajgiri')
    expect(ws.getCell(6, 3).value).toBe(1234567)

    expect(ws.getCell(9, 5).value).toBe(12.35) // Stone Aggregates
    expect(ws.getCell(10, 5).value).toBe(6.79) // Sand
    expect(ws.getCell(11, 5).value).toBe(3.5) // Gravel
    expect(ws.getCell(12, 5).value).toBe(1000) // Granite Slabs
    expect(ws.getCell(13, 5).value).toBe(250.5) // Napa Slabs
    expect(ws.getCell(14, 5).value).toBe(4.44) // Cement
    expect(ws.getCell(15, 5).value).toBe(1.11) // Steel

    // Row/material labels themselves are untouched.
    expect(ws.getCell(9, 2).value).toBe('Stone Aggregates')
    expect(ws.getCell(15, 2).value).toBe('Steel')
  })

  it('leaves header fields untouched when no meta is supplied', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = await fillMaterialTemplate(buffer, totals)
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]
    expect(ws.getCell(3, 3).value).toBeNull()
  })
})

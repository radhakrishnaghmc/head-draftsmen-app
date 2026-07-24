import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import * as ExcelJS from 'exceljs'
import { fillBoqTemplate } from '../core/boqTemplate'

const TEMPLATE_PATH = resolve(__dirname, '../resources/boq-template.xlsx')

describe('fillBoqTemplate', () => {
  it('styles header and item cells Verdana 12pt center/middle, without bleeding into untouched rows', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = await fillBoqTemplate(buffer, [
      { quantity: '10', description: 'Test item one', workType: 'Earth Work', shortDescription: 'Test one', apss: 'NA', rate: '100', uom: 'Cum' },
      { quantity: '20', description: 'Test item two', workType: 'Concrete', shortDescription: 'Test two', apss: 'NA', rate: '200', uom: 'Sqm' }
    ])

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]

    // Header row (1) and both item rows (2, 3) — every column forced to
    // Verdana 12pt, centered/middle.
    for (const r of [1, 2, 3]) {
      for (let c = 1; c <= 8; c++) {
        const cell = ws.getCell(r, c)
        expect(cell.font?.name, `row ${r} col ${c} font`).toBe('Verdana')
        expect(cell.font?.size, `row ${r} col ${c} size`).toBe(12)
        expect(cell.alignment?.horizontal, `row ${r} col ${c} horizontal`).toBe('center')
        expect(cell.alignment?.vertical, `row ${r} col ${c} vertical`).toBe('middle')
      }
    }

    // The total row's Amount cell (row 4, col 8) is also styled...
    const totalAmount = ws.getCell(4, 8)
    expect(totalAmount.font?.name).toBe('Verdana')
    expect(totalAmount.font?.size).toBe(12)

    // ...but the *other* columns of that same row, and every row beyond it,
    // must be untouched — this is the regression guard: ExcelJS shares one
    // style object across every cell using the same style index in the
    // loaded file, so naively mutating `cell.font =` on the header/item
    // cells previously bled the new size/alignment into unrelated rows far
    // below the ones actually being filled.
    for (const r of [4, 5, 6, 8]) {
      for (let c = 1; c <= 7; c++) {
        const cell = ws.getCell(r, c)
        expect(cell.font?.size, `row ${r} col ${c} size should be untouched`).toBe(10)
      }
    }
  }, 30000)

  it('clears the fixed row height on header/item/total rows so Excel auto-fits them', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = await fillBoqTemplate(buffer, [
      { quantity: '10', description: 'Test item one', workType: 'Earth Work', shortDescription: 'Test one', apss: 'NA', rate: '100', uom: 'Cum' },
      { quantity: '20', description: 'Test item two', workType: 'Concrete', shortDescription: 'Test two', apss: 'NA', rate: '200', uom: 'Sqm' }
    ])

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]

    // Header (1), the two item rows (2, 3), and the total row (4) must have
    // no fixed height — the template originally pinned these (customHeight)
    // for its smaller original font, which would clip the now-larger
    // Verdana 12pt wrapped text. Excel only auto-calculates a row's height
    // when it has no explicit height set.
    for (const r of [1, 2, 3, 4]) {
      expect(ws.getRow(r).height, `row ${r} height should be unset for auto-fit`).toBeFalsy()
    }
  }, 30000)

  it('writes the estimate\'s work name two rows below the total (leaving one blank row)', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = await fillBoqTemplate(
      buffer,
      [{ quantity: '10', description: 'Test item one', workType: 'Earth Work', shortDescription: 'Test one', apss: 'NA', rate: '100', uom: 'Cum' }],
      'Road from A to B'
    )

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]

    // header (1) + 1 item row (2) -> total row 3 -> blank row 4 -> name row 5.
    expect(ws.getCell(4, 2).value).toBeFalsy()
    expect(ws.getCell(5, 2).value).toBe('Name of Work: Road from A to B')
  }, 30000)

  it('writes no name row at all when no work name was extracted', async () => {
    const buffer = readFileSync(TEMPLATE_PATH)
    const out = await fillBoqTemplate(buffer, [
      { quantity: '10', description: 'Test item one', workType: 'Earth Work', shortDescription: 'Test one', apss: 'NA', rate: '100', uom: 'Cum' }
    ])

    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(out as unknown as ArrayBuffer)
    const ws = workbook.worksheets[0]
    expect(ws.getCell(5, 2).value).toBeFalsy()
  }, 30000)
})

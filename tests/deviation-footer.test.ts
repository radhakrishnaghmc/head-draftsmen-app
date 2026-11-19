import { readFileSync } from 'fs'
import * as ExcelJS from 'exceljs'
import { expect, it } from 'vitest'
import { fillDeviationTemplate } from '../core/deviationTemplate'

it('prints the circle CNO in the header and signature footer', async () => {
  const items = [{ description: 'x', unit: 'Cum', quantity: '1', rate: '10' }]
  const out = await fillDeviationTemplate(readFileSync('resources/deviation-template.xlsx'), items, { circle: 'Nizampet', circleNumber: '58', estimateAmountLakhs: 5 })
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(out as unknown as ArrayBuffer)
  const texts: string[] = []
  wb.worksheets[0].eachRow((r) => r.eachCell((c) => typeof c.value === 'string' && texts.push(c.value)))
  expect(texts).toContain('NIZAMPET CIRCLE-58')
  expect(new Set(texts.filter((t) => /\nNIZAMPET CIRCLE-58/.test(t))).size).toBe(3)
  expect(wb.worksheets[0].pageSetup.printArea).toMatch(/^A1:M\d+$/)
  expect(texts.some((t) => /^AE\n/.test(t))).toBe(true)
  expect(texts.some((t) => /^SEIGNORAGE CHARGES|^Permit Fee/.test(t))).toBe(false)
  expect(texts.some((t) => /\ncircle-57/i.test(t))).toBe(false)
})

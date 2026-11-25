import { describe, expect, it } from 'vitest'
import * as ExcelJS from 'exceljs'
import { buildMeasurementBookWorkbook } from '../core/measurementBookExcel'
import { emptyMbModel, mergeMbPage } from '../core/measurementBook'

describe('measurement book workbook', () => {
  it('merges page JSON, skips unmatched items in the seigniorage sheet', async () => {
    let m = emptyMbModel()
    m = mergeMbPage(m, 2, '{"header":{"work":"CC road"},"items":[{"description":"RMC M20","unit":"Cum","lines":[{"location":"Rd 4","no":1,"l":14,"b":[5.35,5.35,5.35],"d":[0.15],"written":11.24}],"totalWritten":11.24}],"abstract":[]}')!
    m = mergeMbPage(m, 5, '{"items":[{"continued":true,"lines":[{"location":"Rd 5","no":1,"l":1,"b":[1],"d":[1],"written":1}]}],"abstract":[]}')!
    m = mergeMbPage(m, 9, '{"items":[],"abstract":[{"description":"RMC M20","qtyWritten":11.24,"unit":"Cum","rate":5755.13,"amountWritten":64659,"chartSl":2},{"description":"Groove cutting","qtyWritten":10,"unit":"Rmt","rate":68.17,"amountWritten":681,"chartSl":null}]}')!
    expect(m.items).toHaveLength(1)
    expect(m.items[0].lines).toHaveLength(2)
    expect(mergeMbPage(m, 3, 'not json')).toBeNull()

    const wb = new ExcelJS.Workbook()
    await wb.xlsx.load((await buildMeasurementBookWorkbook(m, 'N-1')) as unknown as ArrayBuffer)
    expect(wb.worksheets.map((w) => w.name)).toEqual(['Measurements', 'Abstract', 'Seigniorage charges'])
    const sg = wb.getWorksheet('Seigniorage charges')!
    const labels: string[] = []
    sg.eachRow((r) => labels.push(String(r.getCell(4).value ?? '')))
    expect(labels.some((l) => l.includes('Groove'))).toBe(false) // no chart item → left out
    expect(labels.some((l) => l.includes('RMC'))).toBe(true)
  })
})

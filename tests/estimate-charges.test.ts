import { describe, it, expect } from 'vitest'
import { extractEstimateCharges } from '../core/estimateExtract'
import { civilTenderPlaceholders, deriveFields } from '../core/workOrderAgreement'

const HEADER = ['S.No', 'Description', 'Qty', 'Rate', 'Per', 'Amount']

function estimateGrid(rows: string[][]): string[][] {
  return [['Name of work:', 'Laying of CC roads'], HEADER, ...rows]
}

describe('extractEstimateCharges', () => {
  it('reads the GST and Labour Cess the estimate charges', () => {
    const grid = estimateGrid([
      ['1', 'Earth work excavation', '100', '250', 'Cum', '25000'],
      ['', 'Sub Total', '', '', '', '1483000'],
      ['', 'Add 1.00% Labour Cess', '', '', '', '14830'],
      ['', 'Sub Total', '', '', '', '1497830'],
      ['', 'Add 18.00% GST', '', '', '', '269609'],
      ['', 'Grand Total', '', '', '', '1800000']
    ])
    expect(extractEstimateCharges(grid, 1)).toEqual({ gst: 269609, labourCess: 14830 })
  })

  it('takes the bottom-most charge rows, so a multi-estimate sheet ends on its abstract', () => {
    const grid = estimateGrid([
      ['', 'Add 18.00% GST', '', '', '', '100000'],
      ['', 'Add 18.00% GST', '', '', '', '269609']
    ])
    expect(extractEstimateCharges(grid, 1).gst).toBe(269609)
  })

  it('ignores a GST registration number in the title block', () => {
    const grid = estimateGrid([['', 'GST No: 36ABCDE1234F1Z5', '', '', '', '']])
    expect(extractEstimateCharges(grid, 1).gst).toBeUndefined()
  })

  it('returns nothing when the estimate charges neither', () => {
    expect(extractEstimateCharges(estimateGrid([['1', 'Earth work', '1', '2', 'Cum', '2']]), 1)).toEqual({})
  })
})

describe('civilTenderPlaceholders GST', () => {
  const fields = deriveFields({}, {}, { 'Name of the work': 'Laying of CC roads', ECV: '1712000' })

  it("prints the estimate's own GST when one was read", () => {
    const v = civilTenderPlaceholders(fields, {}, { gstRupees: 269609 })
    expect(v.GST).toBe('Rs.2,69,609.00')
  })

  it('stays blank when no estimate has been uploaded, even with an ECV', () => {
    const v = civilTenderPlaceholders(fields, {}, {})
    expect(v.GST).toBe('')
  })

  it('keeps Labour Cess at 1% of the ECV', () => {
    const v = civilTenderPlaceholders(fields, {}, { gstRupees: 269609 })
    expect(v['Labour Cess']).toBe('Rs.17,120.00')
  })
})

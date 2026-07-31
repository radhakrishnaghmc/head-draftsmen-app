import { describe, expect, it } from 'vitest'
import { extractElectricalEstimateItems } from '../core/electricalEstimate'

// Mirrors the real CMC electrical estimate: a flat one-row-per-item BOQ whose
// abstract labels ("Sub - Total", "Add: GST …", "Total : Rs") sit in the Qty
// column, with the amount in the Amount column.
const HEADER = ['Sl. No.', 'Descriptiion', 'Qty', 'Rate  Rs.', 'Unit', 'Amount']

describe('extractElectricalEstimateItems', () => {
  it('reads every flat line item and stops at the cost build-up (label in the Qty column)', () => {
    const grid = [
      ['CYBERABAD MUNICIPAL CORPORATION'],
      ['Name of the work: Providing of 11 Nos 9 mtrs poles'],
      HEADER,
      ['1', 'Hot-dip galvanized 9m poles', '11', '18630', 'Each', '204930'],
      ['2', 'Foundation and erection', '11', '9240', 'Each', '101640'],
      ['3', 'Supply of 110W LED Street lights', '44', '2250', 'Each', '99000'],
      ['4', 'Stringing of AB cable', '300', '2.67', 'Mtr', '801'],
      ['', 'Sub - Total', 'Sub - Total', '', '', '405371'],
      ['', 'Add: GST @ 18%', 'Add: GST @ 18%', '', '', '72967'],
      ['', 'Total :  Rs', 'Total :  Rs', '', '', '478338']
    ]
    const items = extractElectricalEstimateItems(grid, 2)
    expect(items.map((i) => ({ d: i.description, q: i.quantity, r: i.rate, u: i.unit }))).toEqual([
      { d: 'Hot-dip galvanized 9m poles', q: '11', r: '18630', u: 'Each' },
      { d: 'Foundation and erection', q: '11', r: '9240', u: 'Each' },
      { d: 'Supply of 110W LED Street lights', q: '44', r: '2250', u: 'Each' },
      { d: 'Stringing of AB cable', q: '300', r: '2.67', u: 'Mtr' }
    ])
    expect(items[0].estimateAmount).toBe('204930')
  })

  it('skips a section-header row (description but no Qty/Rate)', () => {
    const grid = [
      HEADER,
      ['', 'A. POLES', '', '', '', ''],
      ['1', 'Pole', '11', '18630', 'Each', '204930'],
      ['', 'Sub - Total', '', '', '', '204930']
    ]
    const items = extractElectricalEstimateItems(grid, 0)
    expect(items).toHaveLength(1)
    expect(items[0].description).toBe('Pole')
  })

  it('throws when the sheet has no Qty/Rate columns', () => {
    expect(() => extractElectricalEstimateItems([['A', 'B', 'C'], ['x', 'y', 'z']], 0)).toThrow()
  })
})

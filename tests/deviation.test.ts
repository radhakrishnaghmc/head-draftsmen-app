import { describe, expect, it } from 'vitest'
import { extractEstimateAmountLakhs } from '../core/deviation'
import type { EstimateWorkItem } from '../core/estimateExtract'

function item(quantity: string, rate: string): EstimateWorkItem {
  return { description: 'Test item', quantity, rate, unit: 'Cum' }
}

describe('extractEstimateAmountLakhs', () => {
  it('reads an explicit "Estimate Amount" label from the title block', () => {
    const grid = [['Name of Work: Road A'], ['Estimate Amount: 45'], ['S.No', 'Description', 'Qty', 'Rate', 'Unit']]
    expect(extractEstimateAmountLakhs(grid, 2, [item('10', '100')])).toBe(45)
  })

  it('falls back to "ECV" when "Estimate Amount" is not present', () => {
    const grid = [['ECV: 12'], ['S.No', 'Description', 'Qty', 'Rate', 'Unit']]
    expect(extractEstimateAmountLakhs(grid, 1, [item('10', '100')])).toBe(12)
  })

  it('computes from summed item amounts when neither label nor Grand Total is present', () => {
    const grid = [['Some title'], ['S.No', 'Description', 'Qty', 'Rate', 'Unit']]
    // 10*100 + 5*6435.24 = 1000 + 32176.2 = 33176.2 rupees -> 0.331762 Lakhs -> rounds to 0.33.
    const items = [item('10', '100'), item('5', '6435.24')]
    expect(extractEstimateAmountLakhs(grid, 1, items)).toBe(0.33)
  })

  it('uses the Grand Total (not the item-sum/ECV) when no amount label is present', () => {
    // Item-sum is 4.15 Lakhs (= the ECV); the sanctioned Grand Total is 5 Lakhs.
    // Without the Grand Total step the estimate amount would wrongly equal the ECV.
    const grid = [
      ['Some title'],
      ['S.No', 'Description', 'Qty', 'Rate', 'Unit', 'Amount'],
      ['1', 'Work', '4150.2', '100', 'Cum', '415020'],
      ['', 'Grand Total', '', '', '', '500000']
    ]
    expect(extractEstimateAmountLakhs(grid, 1, [item('4150.2', '100')])).toBe(5)
  })
})

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

  it('computes from summed item amounts when neither label is present', () => {
    const grid = [['Some title'], ['S.No', 'Description', 'Qty', 'Rate', 'Unit']]
    // 10*100 + 5*6435.24 = 1000 + 32176.2 = 33176.2 rupees -> 0.331762 Lakhs -> rounds to 0.33.
    const items = [item('10', '100'), item('5', '6435.24')]
    expect(extractEstimateAmountLakhs(grid, 1, items)).toBe(0.33)
  })
})

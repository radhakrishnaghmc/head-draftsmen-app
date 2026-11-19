import { describe, it, expect } from 'vitest'
import { extractEstimateItems } from '../core/estimateExtract'
import { buildBoqFromEstimate, boqToItems } from '../src/boqTransform'

// A flat BOQ has no S.No column: estimate extraction throws, so the upload must fall back to boqToItems.
describe('BOQ upload', () => {
  it('is not an estimate, but boqToItems reads every item', () => {
    const items = [
      { description: 'Earth work excavation', quantity: '10', rate: '100', unit: 'Cum', estimateAmount: '1000' },
      { description: 'PCC 1:4:8', quantity: '5', rate: '200', unit: 'Cum', estimateAmount: '1000' }
    ]
    const boq = buildBoqFromEstimate(items)
    const grid = [boq.headers, ...boq.rows.map((r) => boq.headers.map((h) => r[h] ?? ''))]
    expect(() => extractEstimateItems(grid, 0)).toThrow()
    expect(boqToItems(boq).map((i) => i.description)).toEqual(['Earth work excavation', 'PCC 1:4:8'])
  })
})

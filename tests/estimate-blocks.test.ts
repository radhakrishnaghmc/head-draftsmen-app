import { describe, it, expect } from 'vitest'
import { splitEstimateBlocks, extractGrandTotalLakhs, extractWorkName, extractEstimateItems } from '../core/estimateExtract'

// A minimal estimate block: title, Name of Work, header row, one item, totals.
function estimate(work: string, grandTotal: string): string[][] {
  return [
    ['CYBERABAD MUNICIPAL CORPORATION'],
    [`Name of Work: ${work}`],
    ['Sl. No', 'Description', 'No.s', '', '', 'L', 'B', 'D', 'Qty', 'Rate', 'Per', '', 'Amount'],
    ['1', 'Earth work excavation', '1', 'X', '1', '470', '6', '0.15', '423', '350.28', 'Cum', '', '148168'],
    ['', '', '', '', '', '', '', '', '', 'Total', '', '', '148168'],
    ['', '', '', '', '', '', '', '', '', 'Grand Total', '', '', grandTotal, '0.00']
  ]
}

describe('splitEstimateBlocks', () => {
  it('returns the whole grid unchanged when there is one estimate', () => {
    const grid = estimate('Laying of CC roads at Peddamma nagar', '11200000')
    const blocks = splitEstimateBlocks(grid)
    expect(blocks).toHaveLength(1)
    expect(blocks[0]).toBe(grid)
  })

  it('splits a sheet that stacks several estimates, one per Name of Work', () => {
    const a = estimate('Laying of CC roads at Peddamma nagar', '11200000')
    const b = estimate('Laying of CC roads at Kamalamma nagar', '2000000')
    const blocks = splitEstimateBlocks([...a, ...b])
    expect(blocks).toHaveLength(2)
    // each block starts at its own Name of Work row
    expect(extractWorkName(blocks[0], 1)).toContain('Peddamma')
    expect(extractWorkName(blocks[1], 1)).toContain('Kamalamma')
    expect(extractGrandTotalLakhs(blocks[0], 1)).toBe(112)
    expect(extractGrandTotalLakhs(blocks[1], 1)).toBe(20)
  })

  it('treats a grid with no Name of Work marker as a single block', () => {
    const grid = [['Sl. No', 'Description', 'Amount'], ['1', 'x', '5']]
    expect(splitEstimateBlocks(grid)).toEqual([grid])
  })
})

describe('extractGrandTotalLakhs', () => {
  it('reads the Grand Total from the Amount column, in lakhs', () => {
    const grid = estimate('Some work', '11200000')
    expect(extractGrandTotalLakhs(grid, 2)).toBe(112)
  })

  it('skips a trailing deviation figure and takes the positive amount', () => {
    const grid: string[][] = [
      ['Sl. No', 'Description', 'Amount'],
      ['', 'Grand Total', '5830000', '-1234'] // amount col = index 2
    ]
    expect(extractGrandTotalLakhs(grid, 0)).toBe(58.3)
  })

  it('returns undefined when there is no Grand Total row', () => {
    const grid = [['Sl. No', 'Description', 'Amount'], ['1', 'item', '100']]
    expect(extractGrandTotalLakhs(grid, 0)).toBeUndefined()
  })
})

describe('extractEstimateItems — column-legend row', () => {
  // Departmental templates print a "1 2 3 … 9" column-number row directly under
  // the header; it must not be read as a work item (it once surfaced as a
  // phantom BOQ line with Description "2", Qty "7", Rate "8").
  it('skips the numbered column-legend row under the header', () => {
    const grid: string[][] = [
      ['Sl. No.', 'Description of work', 'No', '', '', 'L', 'B', 'D', 'Qty', 'Rate/per', 'per', 'Amount'],
      ['1', '2', '3', '', '', '4', '5', '6', '7', '8', '', '9'],
      ['1', 'Providing Series lights to the Temples', '', '', '', '', '', '', '884', '30', 'Rmt', '26520'],
      ['2', 'Mic System 400 Watts', '', '', '', '', '', '', '19', '4200', 'Day', '79800'],
      ['', 'Grand Total', '', '', '', '', '', '', '', '', '', '500000']
    ]
    const items = extractEstimateItems(grid, 0)
    expect(items).toHaveLength(2)
    expect(items[0].description).toContain('Series lights')
    expect(items[1].description).toContain('Mic System')
    expect(items.some((it) => it.description === '2')).toBe(false)
  })
})

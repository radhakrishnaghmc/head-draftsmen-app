import { describe, it, expect } from 'vitest'
import {
  splitEstimateBlocks,
  extractGrandTotalLakhs,
  extractWorkName,
  extractEstimateItems,
  looksLikeGeneralAbstractSheet,
  extractAbstractWorkTitle,
  extractAbstractTotalLakhs
} from '../core/estimateExtract'
import { guessHeaderRow } from '../core/sheet'

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

describe('extractEstimateItems — abbreviated serial header', () => {
  // Open-gym / lump-sum "Repairs" estimates label the serial column "S.N."
  // (S for Serial, N for Number) rather than "S.No", and use a two-row header
  // ("MEASUREMENTS"/"Quantity"/"Per" spanning merged cells over an L/B/D +
  // Qty + unit sub-row). "S.N." must still resolve as the serial column so the
  // sheet isn't rejected with "Could not find S.No / Qty / Rate / Unit columns".
  it('resolves an "S.N." serial header on a lump-sum equipment estimate', () => {
    // The serial column is labelled "S.N." (S for Serial, N for Number) rather
    // than "S.No" — the whole sheet was previously rejected with "Could not
    // find S.No / Qty / Rate / Unit columns" because that abbreviation didn't
    // match. It must resolve like any other serial header.
    const grid: string[][] = [
      ['S.N.', 'Description', 'Nos.', '', '', 'L', 'B', 'D', 'Quantity', 'Rate', 'Per', 'Amount'],
      ['1', 'PULL UP STATION', '1', 'x', '1', '', '', '', '1', '60000', 'Nos', '60000'],
      ['2', 'THAICHI SPINNER', '1', 'x', '1', '', '', '', '1', '83850', 'Nos', '83850'],
      ['', 'Grand Total', '', '', '', '', '', '', '', '', '', '143850']
    ]
    const headerRow = guessHeaderRow(grid)
    expect(headerRow).toBe(0)
    const items = extractEstimateItems(grid, headerRow)
    expect(items).toHaveLength(2)
    expect(items[0].description).toContain('PULL UP STATION')
    expect(Number(items[0].quantity)).toBe(1)
    expect(Number(items[0].rate)).toBe(60000)
    expect(items[0].unit).toBe('Nos')
    expect(items[1].description).toContain('THAICHI SPINNER')
  })
})

describe('General Abstract rollup detection', () => {
  // A single sanctioned work split across component sheets ("1) Security
  // Cabin", "2) Store Room", ...), rolled up by one "GENERAL ABSTRACT" sheet
  // that sums them into the final sanctioned total.
  const abstractGrid: string[][] = [
    ['CMC'],
    ['Construction of Sports Arena at Jyothirao Phule ground, Circle-59, CMC.'],
    ['GENERAL ABSTRACT'],
    ['SoR 2021-22'],
    ['Sl.No', 'Description Item', 'Qty', '', 'Amount in Rs.', 'Amount in Lakhs'],
    ['A', 'Working Items'],
    ['1', 'Construction Security Cabin', '', '', '1262478.00', '12.62478'],
    ['2', 'Construction Store Room', '', '', '1238589.00', '12.38589'],
    ['', 'A. Grand-Total of Working Items', '', '', '2501067.00', '25.01067'],
    ['', 'Total (A+B+C)', '', '', '46200000.00', '462']
  ]

  it('recognizes a sheet carrying a GENERAL ABSTRACT heading', () => {
    expect(looksLikeGeneralAbstractSheet(abstractGrid)).toBe(true)
    expect(looksLikeGeneralAbstractSheet([['Sl.No', 'Description', 'Qty', 'Rate', 'Unit']])).toBe(false)
  })

  it('reads the plain-text work title above the heading (no "Name of Work:" label)', () => {
    expect(extractAbstractWorkTitle(abstractGrid)).toContain('Sports Arena')
  })

  it('takes the largest total row as the final sanctioned amount, not the first', () => {
    const headerRow = guessHeaderRow(abstractGrid)
    expect(extractAbstractTotalLakhs(abstractGrid, headerRow)).toBe(462)
  })
})

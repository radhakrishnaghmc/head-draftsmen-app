import { describe, it, expect } from 'vitest'
import { extractRateEntriesFromGrid, buildRateIndex } from '../core/rateDatabase'

describe('extractRateEntriesFromGrid', () => {
  it('extracts a simple Building-Data-style block ending in "Grand Total"', () => {
    const grid = [
      ['Index-code', 'S No', 'Description', 'Unit', 'Quantity', 'Rate Rs.', 'Amount Rs.'],
      ['BLD-CSTN-1-1', '1', 'Cement Mortar (1 : 1)', '', '', '', ''],
      ['', '', 'Cement', 'kg.', '1440.00', '5.10', '7344.00'],
      ['', '', 'Grand Total', '', '', '', '9739.67']
    ]
    const entries = extractRateEntriesFromGrid(grid, 'Building Data')
    expect(entries).toHaveLength(1)
    expect(entries[0]).toMatchObject({ description: 'Cement Mortar (1 : 1)', rate: '9739.67', sheet: 'Building Data' })
    // The breakdown carries the full material/labour buildup for reproducing
    // as a rate-analysis backup sheet in the generated Technical Sanction.
    expect(entries[0].breakdown).toEqual([
      ['BLD-CSTN-1-1', '1', 'Cement Mortar (1 : 1)', '', '', '', ''],
      ['', '', 'Cement', 'kg.', '1440.00', '5.10', '7344.00'],
      ['', '', 'Grand Total', '', '', '', '9739.67']
    ])
  })

  it('prefers the fuller description when the item-start row only has a short title', () => {
    const grid = [
      ['Index-code', 'S. No', '', 'Description', 'Unit', 'Quantity', 'Rate Rs.', 'Amount Rs.'],
      ['RBR-FNDN-1', '1', '', 'Excavation for Structures', '', '', '', ''],
      ['', '', '', 'Earthwork in excavation for structures as per drawing and technical specifications.', '', '', '', ''],
      ['', '', '', 'Mazdoor (Unskilled)', 'day', '3.64', '605.00', '2202.20'],
      ['', '', '', 'Rate per cum = (a+b+c)/10', '', '', '', '308.31']
    ]
    const entries = extractRateEntriesFromGrid(grid, 'RBR-FNDN')
    expect(entries).toHaveLength(1)
    expect(entries[0]).toMatchObject({
      description: 'Earthwork in excavation for structures as per drawing and technical specifications.',
      rate: '308.31',
      sheet: 'RBR-FNDN'
    })
  })

  it('surfaces every rate variant when one description has several total/rate lines', () => {
    const grid = [
      ['Index-code', 'S. No', '', 'Description', 'Unit', 'Quantity', 'Rate Rs.', 'Amount Rs.'],
      ['RBR-FNDN-1', '1', '', 'Excavation for Structures', '', '', '', ''],
      ['', '', '', 'Earthwork in excavation for structures.', '', '', '', ''],
      ['', '', '(A)', 'Manual Means', '', '', '', ''],
      ['', '', '', 'Cost for 10 cum = a+b+c', '', '', '', '3083.08'],
      ['', '', '', 'Rate per cum = (a+b+c)/10', '', '', '', '308.31'],
      ['', '', '(B)', 'Mechanical Means', '', '', '', ''],
      ['', '', '', 'Cost for 240 cum = a+b+c+d', '', '', '', '33893.94'],
      ['', '', '', 'Rate per cum = (a+b+c+d)/240', '', '', '', '141.22'],
      ['RBR-FNDN-2', '2', '', 'Next item', '', '', '', '']
    ]
    const entries = extractRateEntriesFromGrid(grid, 'RBR-FNDN')
    const excavation = entries.filter((e) => e.description === 'Earthwork in excavation for structures.')
    expect(excavation.map((e) => e.rate)).toEqual(['308.31', '141.22'])
    // The second variant's breakdown starts fresh after the first variant's
    // rate line — it must not carry the first variant's rows along with it.
    expect(excavation[1].breakdown.map((r) => r[2] || r[3])).toEqual([
      '(B)',
      'Cost for 240 cum = a+b+c+d',
      'Rate per cum = (a+b+c+d)/240'
    ])
    expect(excavation[1].breakdown.some((r) => r.includes('(A)'))).toBe(false)
  })

  it('picks the numeric value closest to the label, not a stray reference number further right', () => {
    const grid = [
      ['S.No', 'Description'],
      ['1', 'Plastering 12mm thick in two coats'],
      ['', 'Rate per 1 sqm', '', '', '', '593.72', '', '', '', '', '483']
    ]
    const entries = extractRateEntriesFromGrid(grid, 'worked out data')
    expect(entries).toHaveLength(1)
    expect(entries[0]).toMatchObject({ description: 'Plastering 12mm thick in two coats', rate: '593.72', sheet: 'worked out data' })
  })

  it('yields nothing for a sheet with no recognizable item-start convention', () => {
    const grid = [
      ['Sl.no', 'Labour', 'SSR 2022-23', 'Rate'],
      ['1', 'I st class Mason', 'x', '680']
    ]
    expect(extractRateEntriesFromGrid(grid, 'Labour & material')).toEqual([])
  })
})

describe('buildRateIndex', () => {
  it('flattens entries across multiple sheets', () => {
    const sheets = [
      { name: 'A', grid: [['1', 'Item one', '', '', ''], ['', '', 'Total', '', '100']] },
      { name: 'B', grid: [['1', 'Item two', '', '', ''], ['', '', 'Total', '', '200']] }
    ]
    const index = buildRateIndex(sheets)
    expect(index.map((e) => ({ description: e.description, rate: e.rate, sheet: e.sheet }))).toEqual([
      { description: 'Item one', rate: '100', sheet: 'A' },
      { description: 'Item two', rate: '200', sheet: 'B' }
    ])
  })
})

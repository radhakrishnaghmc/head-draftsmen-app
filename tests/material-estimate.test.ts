import { describe, expect, it } from 'vitest'
import { computeMaterialTotals } from '../core/materialEstimate'
import type { EstimateWorkItem } from '../core/estimateExtract'

function item(description: string, quantity: string, unit: string, rate = '100'): EstimateWorkItem {
  return { description, quantity, rate, unit }
}

describe('computeMaterialTotals', () => {
  it('splits a nominal-mix (ratio) concrete item into cement/sand/aggregate via the 1.54 dry-volume factor', () => {
    const { totals, matchedItemCount } = computeMaterialTotals([
      item('Providing PCC 1:2:4 for foundation', '10', 'Cum')
    ])
    // dryVol = 10*1.54 = 15.4; sum=7; cement=15.4*1/7=2.2 cum *1.44 = 3.168 MT
    expect(totals.cementMt).toBeCloseTo(3.168, 2)
    expect(totals.sandMt).toBeCloseTo((15.4 * 2 / 7) * 1.6, 2)
    expect(totals.stoneAggregatesMt).toBeCloseTo((15.4 * 4 / 7) * 1.55, 2)
    expect(matchedItemCount).toBe(1)
  })

  it('maps a nominal concrete grade (M20) to its traditional 1:1.5:3 ratio', () => {
    const { totals } = computeMaterialTotals([item('RCC M20 for roof slab', '5', 'Cum')])
    expect(totals.cementMt).toBeGreaterThan(0)
    expect(totals.sandMt).toBeGreaterThan(0)
    expect(totals.stoneAggregatesMt).toBeGreaterThan(0)
  })

  it('falls back to the design-mix default for a design-mix grade (M25+) with no fixed ratio', () => {
    const { totals } = computeMaterialTotals([item('RCC M30 for column', '4', 'Cum')])
    expect(totals.cementMt).toBeCloseTo(4 * 0.34, 5)
  })

  it('leaves a bare "concrete" item with neither ratio nor grade unmatched, rather than guessing a mix', () => {
    const { totals, matchedItemCount } = computeMaterialTotals([item('Concrete work as directed', '3', 'Cum')])
    expect(matchedItemCount).toBe(0)
    expect(totals.cementMt).toBe(0)
  })

  it('computes brick masonry mortar (cement+sand only, no aggregate) from a 2-part ratio', () => {
    const { totals } = computeMaterialTotals([item('Brick work in CM 1:6 for walls', '20', 'Cum')])
    expect(totals.cementMt).toBeGreaterThan(0)
    expect(totals.sandMt).toBeGreaterThan(0)
    expect(totals.stoneAggregatesMt).toBe(0)
  })

  it('computes plaster mortar from area, ratio, and thickness (default 12mm)', () => {
    const { totals } = computeMaterialTotals([item('12mm cement plastering CM 1:6 to walls', '100', 'Sqm')])
    expect(totals.cementMt).toBeGreaterThan(0)
    expect(totals.sandMt).toBeGreaterThan(0)
  })

  it('converts granite and napa flooring area (Sqm) to Sq.Ft, keeping them separate', () => {
    const { totals } = computeMaterialTotals([
      item('Granite flooring 18mm thick', '10', 'Sqm'),
      item('Napa slab flooring', '5', 'Sqm')
    ])
    expect(totals.graniteSqft).toBeCloseTo(10 * 10.7639, 1)
    expect(totals.napaSqft).toBeCloseTo(5 * 10.7639, 1)
  })

  it('routes WBM/gravel filling to Gravel, not Stone Aggregates, even though it names "aggregate"', () => {
    const { totals } = computeMaterialTotals([item('Water bound macadam using crushable aggregates', '8', 'Cum')])
    expect(totals.gravelMt).toBeCloseTo(8 * 1.75, 5)
    expect(totals.stoneAggregatesMt).toBe(0)
  })

  it('routes Wet Mix Macadam (WMM) to Gravel, not Stone Aggregates, matching real road-estimate wording', () => {
    const { totals, matchedItemCount } = computeMaterialTotals([
      item(
        'Providing, laying, spreading and compacting graded stone aggregate to wet mix macadam specification including premixing the material with water at OMC',
        '36',
        'cum'
      )
    ])
    expect(totals.gravelMt).toBeCloseTo(36 * 1.75, 5)
    expect(totals.stoneAggregatesMt).toBe(0)
    expect(matchedItemCount).toBe(1)
  })

  it('credits Dense Bituminous Macadam (DBM) toward Stone Aggregates without touching cement/sand', () => {
    const { totals } = computeMaterialTotals([
      item(
        'Providing and laying dense graded bituminous macadam with 100-120 TPH batch type HMP using crushed aggregates premixed with bituminous binder VG-30 @ 4.5 per cent by weight',
        '186',
        'cum'
      )
    ])
    expect(totals.stoneAggregatesMt).toBeGreaterThan(0)
    expect(totals.cementMt).toBe(0)
    expect(totals.sandMt).toBe(0)
  })

  it('credits Bituminous Concrete toward Stone Aggregates instead of falling into the cement-concrete ratio rule', () => {
    const { totals } = computeMaterialTotals([
      item(
        'Providing and laying bituminous concrete with 100-120 TPH batch type hot mix plant using crushed aggregates premixed with bituminous binder VG-30 @ 5.4 per cent of mix and filler',
        '111.6',
        'cum'
      )
    ])
    expect(totals.stoneAggregatesMt).toBeCloseTo(111.6 * 2.35 * 0.95, 2)
    expect(totals.cementMt).toBe(0)
  })

  it('leaves a thin bitumen primer/tack coat (Sqm, no volume) unmatched rather than guessing an aggregate quantity', () => {
    const { totals, matchedItemCount } = computeMaterialTotals([
      item('Providing and applying tack coat with Bitumen emulsion (RS-1) using emulsion distributor', '3720', 'sqm')
    ])
    expect(matchedItemCount).toBe(0)
    expect(totals.stoneAggregatesMt).toBe(0)
  })

  it('sums a standalone reinforcement steel item given directly in Kg', () => {
    const { totals } = computeMaterialTotals([item('Providing TMT reinforcement Fe500', '2500', 'Kg')])
    expect(totals.steelMt).toBeCloseTo(2.5, 5)
  })

  it('applies a Data Sheet breakdown coefficient (per-unit) scaled by the item quantity, skipping labour rows', () => {
    const items = [item('PCC 1:2:4 for foundation', '10', 'Cum')]
    const breakdown = [
      ['Cement', '0.22', 'Cum', '350', '77'],
      ['Sand', '0.44', 'Cum', '900', '396'],
      ['Coarse Aggregate', '0.88', 'Cum', '850', '748'],
      ['Cement Mason', '0.5', 'Day', '700', '350']
    ]
    const { totals, matchedItemCount } = computeMaterialTotals(items, [breakdown])
    expect(totals.cementMt).toBeCloseTo(10 * 0.22 * 1.44, 5)
    expect(totals.sandMt).toBeCloseTo(10 * 0.44 * 1.6, 5)
    expect(totals.stoneAggregatesMt).toBeCloseTo(10 * 0.88 * 1.55, 5)
    expect(matchedItemCount).toBe(1)
  })

  it('falls back to standard coefficients when the Data Sheet breakdown has no recognizable material row', () => {
    const items = [item('PCC 1:2:4 for foundation', '10', 'Cum')]
    const breakdown = [['Mason', '0.5', 'Day', '700', '350']]
    const { totals } = computeMaterialTotals(items, [breakdown])
    expect(totals.cementMt).toBeGreaterThan(0) // came from the generic ratio rule instead
  })
})

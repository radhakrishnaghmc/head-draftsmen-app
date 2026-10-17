import { describe, it, expect } from 'vitest'
import { buildMaterialRag, matchMaterialRag } from '../core/materialQuantityRag'
import type { MaterialRag } from '../core/materialQuantityRag'
import { computeMaterialTotals } from '../core/materialEstimate'
import type { RateEntry } from '../core/rateDatabase'
import type { EstimateWorkItem } from '../core/estimateExtract'
import bundled from '../core/data/materialQuantityRag.json'

const item = (description: string, quantity: string, unit: string): EstimateWorkItem =>
  ({ description, quantity, unit, rate: '1' }) as EstimateWorkItem

const entry = (description: string, code: string, unit: string, rows: string[][]): RateEntry => ({
  description,
  rate: '5000',
  sheet: 'Building Data',
  code,
  unit,
  breakdown: [[`${code} | ${description}`], ['Unit = 1cum'], ['A. MATERIALS:'], ...rows]
})

describe('buildMaterialRag', () => {
  const rag = buildMaterialRag(
    [
      entry('Brick masonry in CM 1:8', 'BLD-CSTN-5-5', 'cum', [
        ['Cement', 'kg', '36.00', '5.10', '183.60'],
        ['Fine aggregate (Sand)', 'cum', '0.10', '3477.43', '347.74'],
        ['Seigniorage charges F.A', 'cum', '0.20'],
        ['Mason 1st class', 'day', '0.24', '795.00', '190.80']
      ]),
      // A one-row entry is a flat rate list: its number is a price, not a quantity.
      { description: 'High Polished Granite', rate: '2789', sheet: 'Labour & material', unit: 'sqm', breakdown: [['Granite', 'sqm', '2789', '1']] }
    ],
    'test'
  )

  it('keeps per-unit material lines and drops labour, charge rows and flat rate lists', () => {
    expect(rag.entries).toHaveLength(1)
    expect(rag.entries[0].materials.map((m) => [m.material, m.unit, m.amount])).toEqual([
      ['cement', 'kg', 36],
      ['sand', 'cum', 0.1]
    ])
  })

  it('lists, per material, every item that uses it', () => {
    expect(rag.byMaterial.cement).toEqual([0])
    expect(rag.byMaterial.sand).toEqual([0])
    expect(rag.byMaterial.steel).toEqual([])
    expect(rag.byMaterial.granite).toEqual([])
  })

  it('does not store rates', () => {
    expect(JSON.stringify(rag)).not.toContain('5000')
  })
})

describe('matchMaterialRag + computeMaterialTotals', () => {
  const rag: MaterialRag = buildMaterialRag(
    [
      entry('Brick masonry in CM 1:8 with traditional bricks', 'BLD-CSTN-5-5', 'cum', [
        ['Cement', 'kg', '36.00', '5.10', '183.60'],
        ['Fine aggregate (Sand)', 'cum', '0.10', '3477.43', '347.74']
      ])
    ],
    'test'
  )

  it('matches an estimate item by Data Sheet code and scales the quantities by the item quantity', () => {
    const items = [item('BLD-CSTN-5-5: Brick masonry in CM 1:8', '20', 'cum')]
    const lines = matchMaterialRag(items, rag)
    expect(lines[0]).toBeDefined()
    const { totals } = computeMaterialTotals(items, undefined, lines)
    expect(totals.cementMt).toBeCloseTo((36 * 20) / 1000, 6) // 0.72 MT
    expect(totals.sandMt).toBeCloseTo(0.1 * 20 * 1.6, 6)
  })

  it('leaves an unmatched item to the standard rules', () => {
    const items = [item('Supply and fixing TMT Fe 500 reinforcement', '2500', 'kg')]
    const lines = matchMaterialRag(items, rag)
    expect(lines[0]).toBeUndefined()
    expect(computeMaterialTotals(items, undefined, lines).totals.steelMt).toBeCloseTo(2.5, 6)
  })
})

describe('bundled material-quantity index', () => {
  const rag = bundled as unknown as MaterialRag

  it('covers every material the Material Quantity sheet reports', () => {
    for (const k of ['cement', 'steel', 'sand', 'aggregate'] as const) expect(rag.byMaterial[k].length).toBeGreaterThan(0)
  })

  it('indexes only entries that really carry materials', () => {
    expect(rag.entries.length).toBeGreaterThan(50)
    for (const e of rag.entries) expect(e.materials.length).toBeGreaterThan(0)
  })
})

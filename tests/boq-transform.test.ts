import { describe, it, expect } from 'vitest'
import {
  boqToScheduleA,
  computeEcvFromItems,
  extractWorkNameFromBoq,
  boqColumnsMatchedViaEmbedding,
  BOQ_SCHEDULE_A_COLUMN_SPECS
} from '../src/boqTransform'
import type { ExcelTable } from '../core/types'
import type { EstimateWorkItem } from '../core/estimateExtract'

function table(headers: string[], rows: Record<string, string>[]): ExcelTable {
  return { id: 'boq', name: 'BOQ', path: '/tmp/boq.xlsx', headers, rows }
}

describe('boqToScheduleA', () => {
  it('finds the quantity column by its full name', () => {
    const boq = table(
      ['S.No', 'Quantity', 'Description of Item', 'Rate', 'Unit', 'Amount'],
      [{ 'S.No': '1', Quantity: '10', 'Description of Item': 'Earth work', Rate: '100', Unit: 'Cum', Amount: '1000' }]
    )
    const result = boqToScheduleA(boq)
    expect(result.rows[0]['Probable Quantity']).toBe('10')
  })

  it('finds the quantity column via its "QTY" alias', () => {
    const boq = table(
      ['S.No', 'QTY', 'Description of Item', 'Rate', 'Unit', 'Amount'],
      [{ 'S.No': '1', QTY: '25', 'Description of Item': 'Concreting', Rate: '200', Unit: 'Cum', Amount: '5000' }]
    )
    const result = boqToScheduleA(boq)
    expect(result.rows[0]['Probable Quantity']).toBe('25')
    expect(result.rows[0]['Description of item']).toBe('Concreting')
    expect(result.rows[0]['Estimated Rate in Rs. & Ps.']).toBe('200')
    expect(result.rows[0].Units).toBe('Cum')
    expect(result.rows[0]['Estimated Amount in Rs']).toBe('5000')
  })

  it('finds description/unit columns via their common aliases', () => {
    const boq = table(
      ['Sl No', 'Qty', 'Item Description', 'Rate', 'UOM', 'Amount'],
      [{ 'Sl No': '1', Qty: '5', 'Item Description': 'Brick work', Rate: '50', UOM: 'Sqm', Amount: '250' }]
    )
    const result = boqToScheduleA(boq)
    expect(result.rows[0]['Probable Quantity']).toBe('5')
    expect(result.rows[0]['Description of item']).toBe('Brick work')
    expect(result.rows[0].Units).toBe('Sqm')
  })

  it('untangles a source file that has Rate and UOM data swapped under their headers', () => {
    const boq = table(
      ['Estimate Quantity (only Figures)', 'Item Detailed \nSpecification Description', 'Rate (INR)', 'UOM', 'Amount (INR)'],
      [
        {
          'Estimate Quantity (only Figures)': '237.06',
          'Item Detailed \nSpecification Description': 'Earth work excavation',
          'Rate (INR)': 'Cum',
          UOM: '308.31',
          'Amount (INR)': '73088.00'
        },
        {
          'Estimate Quantity (only Figures)': '90.00',
          'Item Detailed \nSpecification Description': 'Drilling of tube wells',
          'Rate (INR)': 'Rmt',
          UOM: '298.00',
          'Amount (INR)': '26820.00'
        }
      ]
    )
    const result = boqToScheduleA(boq)
    expect(result.rows[0]['Estimated Rate in Rs. & Ps.']).toBe('308.31')
    expect(result.rows[0].Units).toBe('Cum')
    expect(result.rows[1]['Estimated Rate in Rs. & Ps.']).toBe('298.00')
    expect(result.rows[1].Units).toBe('Rmt')
  })

  it('drops a stray note row that has a description but no quantity or amount', () => {
    const boq = table(
      ['Estimate Quantity (only Figures)', 'Item Detailed \nSpecification Description', 'Rate (INR)', 'UOM', 'Amount (INR)'],
      [
        {
          'Estimate Quantity (only Figures)': '10',
          'Item Detailed \nSpecification Description': 'Earth work',
          'Rate (INR)': '100',
          UOM: 'Cum',
          'Amount (INR)': '1000'
        },
        {
          'Estimate Quantity (only Figures)': '',
          'Item Detailed \nSpecification Description': 'Drilling of Injection borewell at KGR Convention',
          'Rate (INR)': '',
          UOM: '',
          'Amount (INR)': ''
        }
      ]
    )
    const result = boqToScheduleA(boq)
    expect(result.rows).toHaveLength(1)
    expect(result.rows[0]['Description of item']).toBe('Earth work')
  })

  it('resolves a Quantity column no regex would recognize, via the embedding fallback', () => {
    // "Nos." matches none of BOQ_SCHEDULE_A_COLUMN_SPECS' Quantity patterns.
    const headers = ['S.No', 'Nos.', 'Description of Item', 'Rate', 'Unit', 'Amount']
    const boq = table(headers, [
      { 'S.No': '1', 'Nos.': '10', 'Description of Item': 'Earth work', Rate: '100', Unit: 'Cum', Amount: '1000' }
    ])

    expect(() => boqToScheduleA(boq)).toThrow(/"Estimate Quantity"/)

    // One embedding per header, then one per spec (in BOQ_SCHEDULE_A_COLUMN_SPECS order:
    // Estimate Quantity, Item Detailed Specification Description, Rate, UOM, Amount) —
    // "Nos." is deliberately closest to the Estimate Quantity label's vector.
    const zero = [0, 0, 0, 0]
    const embeddings = {
      headerVectors: [zero, [1, 0, 0, 0], zero, zero, zero, zero],
      labelVectors: BOQ_SCHEDULE_A_COLUMN_SPECS.map((s) => (s.label === 'Estimate Quantity' ? [1, 0, 0, 0] : zero))
    }

    const result = boqToScheduleA(boq, embeddings)
    expect(result.rows[0]['Probable Quantity']).toBe('10')
    expect(boqColumnsMatchedViaEmbedding(headers, embeddings)).toEqual(['Estimate Quantity'])
  })
})

describe('computeEcvFromItems', () => {
  function item(quantity: string, rate: string): EstimateWorkItem {
    return { description: 'Test item', quantity, rate, unit: 'Cum' }
  }

  it('sums each item\'s quantity × rate, rounded per item, matching the BOQ template\'s own formula', () => {
    // 237.06 * 308.31 = 73,088.0... -> rounds to 73088; 90 * 298 = 26820 exactly.
    const total = computeEcvFromItems([item('237.06', '308.31'), item('90', '298')])
    expect(total).toBe(73088 + 26820)
  })

  it('treats unparseable quantity/rate as 0 rather than throwing', () => {
    expect(computeEcvFromItems([item('', '100'), item('10', '')])).toBe(0)
  })
})

describe('extractWorkNameFromBoq', () => {
  it('finds the "Name of Work: ..." marker regardless of which column it landed in', () => {
    const boq = table(
      ['Estimate Quantity', 'Item Detailed Specification Description', 'Rate', 'UOM', 'Amount'],
      [
        { 'Estimate Quantity': '10', 'Item Detailed Specification Description': 'Earth work', Rate: '100', UOM: 'Cum', Amount: '1000' },
        { 'Estimate Quantity': '', 'Item Detailed Specification Description': '', Rate: '', UOM: '', Amount: '1000' },
        { 'Estimate Quantity': '', 'Item Detailed Specification Description': 'Name of Work: Road from A to B', Rate: '', UOM: '', Amount: '' }
      ]
    )
    expect(extractWorkNameFromBoq(boq)).toBe('Road from A to B')
  })

  it('returns undefined when the BOQ has no such marker', () => {
    const boq = table(
      ['Estimate Quantity', 'Item Detailed Specification Description', 'Rate', 'UOM', 'Amount'],
      [{ 'Estimate Quantity': '10', 'Item Detailed Specification Description': 'Earth work', Rate: '100', UOM: 'Cum', Amount: '1000' }]
    )
    expect(extractWorkNameFromBoq(boq)).toBeUndefined()
  })
})

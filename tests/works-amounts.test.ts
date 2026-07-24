import { describe, expect, it } from 'vitest'
import {
  lakhsToRupees,
  indianDigitGroups,
  formatRupees,
  computeWorkAmounts,
  withComputedAmounts,
  rupeesToLakhsString,
  applyEcvFromBoq
} from '../core/worksAmounts'
import type { ExcelTable } from '../core/types'

describe('lakhsToRupees', () => {
  it('converts Lakhs to rupees', () => {
    expect(lakhsToRupees('1')).toBe(100000)
    expect(lakhsToRupees('.30')).toBe(30000)
    expect(lakhsToRupees('45')).toBe(4500000)
  })
})

describe('indianDigitGroups', () => {
  it('groups digits the Indian way (last 3, then pairs)', () => {
    expect(indianDigitGroups(100000)).toBe('1,00,000')
    expect(indianDigitGroups(4500000)).toBe('45,00,000')
    expect(indianDigitGroups(999)).toBe('999')
    expect(indianDigitGroups(1234567)).toBe('12,34,567')
  })
})

describe('formatRupees', () => {
  it('wraps the Indian-grouped figure in Rs .../-', () => {
    expect(formatRupees(100000)).toBe('Rs 1,00,000/-')
  })
})

describe('computeWorkAmounts', () => {
  it('computes EMD @ 1%/1.5% from ECV, falling back to the estimate when ECV is blank', () => {
    const c = computeWorkAmounts({ 'Amount of estimate': '45', 'Estimate Amount ECV': '' })
    expect(c.estimate).toBe(4500000)
    expect(c.ecv).toBe(4500000) // fell back to the estimate
    expect(c.emd1).toBe(45000)
    expect(c.emd1_5).toBe(67500)
  })

  it('uses ECV over the estimate when both are present', () => {
    const c = computeWorkAmounts({ 'Amount of estimate': '10', 'Estimate Amount ECV': '12' })
    expect(c.estimate).toBe(1000000)
    expect(c.ecv).toBe(1200000)
    expect(c.emd1).toBe(12000)
    expect(c.emd1_5).toBe(18000)
  })

  it('charges ASD only once Tender Percentage exceeds 25%, at (percentage - 25%) of ECV', () => {
    const under = computeWorkAmounts({ 'Estimate Amount ECV': '10', 'Tender Percentage': '20' })
    expect(under.asd).toBe(0)

    const over = computeWorkAmounts({ 'Estimate Amount ECV': '10', 'Tender Percentage': '30' })
    // 30% - 25% = 5% of 1,000,000 = 50,000
    expect(over.asd).toBe(50000)
  })

  it('computes Contract Amount as ECV net of the tendered percentage', () => {
    const c = computeWorkAmounts({ 'Estimate Amount ECV': '10', 'Tender Percentage': '18' })
    // 1,000,000 * (1 - 0.18) = 820,000
    expect(c.contractAmount).toBe(820000)
  })

  it('leaves Contract Amount null when Tender Percentage is not available, rather than assuming 0%', () => {
    const missing = computeWorkAmounts({ 'Estimate Amount ECV': '10' })
    expect(missing.contractAmount).toBeNull()

    const blank = computeWorkAmounts({ 'Estimate Amount ECV': '10', 'Tender Percentage': '' })
    expect(blank.contractAmount).toBeNull()

    // A genuine 0% tender, though, is a real value and should compute (equal to ECV).
    const zero = computeWorkAmounts({ 'Estimate Amount ECV': '10', 'Tender Percentage': '0' })
    expect(zero.contractAmount).toBe(1000000)
  })
})

describe('withComputedAmounts', () => {
  it('replaces every amount column with its computed, Rs-formatted value', () => {
    const row = withComputedAmounts({
      'Amount of estimate': '45',
      'Estimate Amount ECV': '',
      'Tender Percentage': '30'
    })
    expect(row['Amount of estimate']).toBe('Rs 45,00,000/-')
    expect(row['Estimate Amount']).toBe('Rs 45,00,000/-')
    expect(row['Estimate Amount ECV']).toBe('Rs 45,00,000/-')
    expect(row['EMD 1%']).toBe('Rs 45,000/-')
    expect(row['EMD 1.5%']).toBe('Rs 67,500/-')
    expect(row.ASD).toBe('Rs 2,25,000/-') // (30% - 25%) of 45,00,000
    expect(row['Contract Amount']).toBe('Rs 31,50,000/-') // 45,00,000 * (1 - 0.30)
  })

  it('leaves Contract Amount blank when Tender Percentage is not available', () => {
    const row = withComputedAmounts({ 'Amount of estimate': '45', 'Estimate Amount ECV': '' })
    expect(row['Contract Amount']).toBe('')
  })
})

describe('rupeesToLakhsString', () => {
  it('formats a rupee figure as a plain Lakhs decimal, matching manual Works List entries', () => {
    expect(rupeesToLakhsString(2500000)).toBe('25')
    expect(rupeesToLakhsString(320000)).toBe('3.2')
    expect(rupeesToLakhsString(45000)).toBe('0.45')
  })
})

describe('applyEcvFromBoq', () => {
  const table: ExcelTable = {
    id: 't1',
    name: 'Works database',
    path: '',
    headers: ['Name of the work', 'Estimate Amount ECV', 'EMD 1%', 'EMD 1.5%'],
    rows: [
      { 'Name of the work': 'Road from A to B', 'Estimate Amount ECV': '', 'EMD 1%': '', 'EMD 1.5%': '' },
      { 'Name of the work': 'Bridge over river', 'Estimate Amount ECV': '', 'EMD 1%': '', 'EMD 1.5%': '' }
    ]
  }

  it('matches the row by name (case/whitespace-insensitive) and fills ECV + EMD @ 1%/1.5%, in Lakhs', () => {
    const { table: out, matched } = applyEcvFromBoq(table, '  road   from a to b ', 2500000)
    expect(matched).toBe(true)
    expect(out.rows[0]['Estimate Amount ECV']).toBe('25')
    expect(out.rows[0]['EMD 1%']).toBe('0.25')
    expect(out.rows[0]['EMD 1.5%']).toBe('0.375')
    // The other row is untouched.
    expect(out.rows[1]['Estimate Amount ECV']).toBe('')
  })

  it('returns matched: false, table unchanged, when no row name matches', () => {
    const result = applyEcvFromBoq(table, 'Some unrelated work', 2500000)
    expect(result.matched).toBe(false)
    expect(result.table).toBe(table)
  })

  it('falls back to the closest embedding match when the exact name differs, flagging matchedViaAi', () => {
    const result = applyEcvFromBoq(table, 'Road works between A and B, Ph-1', 2500000, {
      workNameVector: [1, 0],
      rowNameVectors: [
        [0.99, 0.01], // row 0 ("Road from A to B") — closest
        [0, 1] // row 1 ("Bridge over river") — unrelated
      ]
    })
    expect(result.matched).toBe(true)
    expect(result.matchedViaAi).toBe(true)
    expect(result.table.rows[0]['Estimate Amount ECV']).toBe('25')
  })

  it('does not use an embedding match below the threshold', () => {
    const result = applyEcvFromBoq(table, 'Completely unrelated text', 2500000, {
      workNameVector: [1, 0],
      rowNameVectors: [
        [0.1, 0.99],
        [0.2, 0.98]
      ]
    })
    expect(result.matched).toBe(false)
    expect(result.matchedViaAi).toBeUndefined()
  })
})

import { describe, it, expect } from 'vitest'
import { findWincodeUniquenessViolations } from '../core/worksListAgent'
import type { ExcelTable } from '../core/types'

function table(rows: Record<string, string>[]): ExcelTable {
  return { id: 't', name: 'Works database', path: '', headers: ['Wincode', 'Name of the work'], rows }
}

describe('findWincodeUniquenessViolations', () => {
  it('finds nothing wrong in a clean list', () => {
    const t = table([
      { Wincode: 'D058-1', 'Name of the work': 'Laying of CC road A' },
      { Wincode: 'D058-2', 'Name of the work': 'Laying of CC road B' }
    ])
    expect(findWincodeUniquenessViolations(t)).toEqual([])
  })

  it('flags a real bug found in the Nizampet Works List: one Wincode used for two unrelated works', () => {
    // Real: state-nizampet.json, CMC|Quthbullapur|Nizampet, Wincode D058-26014491
    // shared by Tender ID 715305 (Lavanya Residency to Vasanthi Nilay...) and
    // Tender ID 710624 (Harithavanam colony Main Road) — two different roads.
    const t = table([
      { Wincode: 'D058-26014491', 'Name of the work': 'laying of CC road from Lavanya Residency to Vasanthi Nilay' },
      { Wincode: 'D058-26014491', 'Name of the work': 'Laying of CC Road in Harithavanam colony Main Road' }
    ])
    const violations = findWincodeUniquenessViolations(t)
    expect(violations).toHaveLength(1)
    expect(violations[0]).toMatchObject({ type: 'duplicate-wincode', key: 'D058-26014491', rowIndices: [0, 1] })
  })

  it('does NOT flag the same work legitimately re-entered under the same Wincode with a recall/reservation tag added', () => {
    const t = table([
      { Wincode: 'D058-1', 'Name of the work': 'Laying of CC road at Plot No 17 (Reserved for SC)' },
      { Wincode: 'D058-1', 'Name of the work': 'Laying of CC road at Plot No 17' }
    ])
    expect(findWincodeUniquenessViolations(t)).toEqual([])
  })

  it('flags one work name entered under two different Wincodes', () => {
    const t = table([
      { Wincode: 'D058-1', 'Name of the work': 'Laying of CC road at Plot No 17' },
      { Wincode: 'D058-2', 'Name of the work': 'Laying of CC road at Plot No 17' }
    ])
    const violations = findWincodeUniquenessViolations(t)
    expect(violations).toHaveLength(1)
    expect(violations[0]).toMatchObject({
      type: 'duplicate-name',
      key: 'Laying of CC road at Plot No 17',
      rowIndices: [0, 1]
    })
  })

  it('never flags blank Wincodes or blank names against each other', () => {
    const t = table([
      { Wincode: '', 'Name of the work': 'Work with no Wincode yet' },
      { Wincode: '', 'Name of the work': 'Another work with no Wincode yet' },
      { Wincode: 'D058-1', 'Name of the work': '' },
      { Wincode: 'D058-2', 'Name of the work': '' }
    ])
    expect(findWincodeUniquenessViolations(t)).toEqual([])
  })

  it('handles three+ rows colliding on the same Wincode as one violation, not several', () => {
    const t = table([
      { Wincode: 'D058-1', 'Name of the work': 'Work A' },
      { Wincode: 'D058-1', 'Name of the work': 'Work B' },
      { Wincode: 'D058-1', 'Name of the work': 'Work C' }
    ])
    const violations = findWincodeUniquenessViolations(t)
    expect(violations).toHaveLength(1)
    expect(violations[0].rowIndices).toEqual([0, 1, 2])
  })
})

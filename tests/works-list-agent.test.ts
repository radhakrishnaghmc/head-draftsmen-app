import { describe, it, expect } from 'vitest'
import {
  findWincodeUniquenessViolations,
  findMonitoringFormatMismatches,
  findMonitoringFormatWorksIdentityViolations
} from '../core/worksListAgent'
import type { ExcelTable } from '../core/types'
import type { MonitoringFormatSummary, MonitoringFormatWorkRow } from '../core/monitoringFormat'

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

function zeroBucket() {
  return { no: 0, amt: 0 }
}

function mfRow(itemType: string, no: number, amt: number): MonitoringFormatSummary['rows'][number] {
  return {
    itemType,
    totalWorks: { no, amt },
    completed: zeroBucket(),
    upto25: zeroBucket(),
    upto50: zeroBucket(),
    upto75: zeroBucket(),
    above75: zeroBucket(),
    progressTotal: zeroBucket(),
    toBeStarted: zeroBucket(),
    tenderProcess: zeroBucket(),
    heldUp: zeroBucket(),
    cancelled: zeroBucket()
  }
}

function summary(rows: MonitoringFormatSummary['rows'], wardRows?: MonitoringFormatSummary['rows']): MonitoringFormatSummary {
  return {
    officeLabel: '58-Nizampet',
    sheetName: 'C58 MF',
    asOfDate: '23.07.2026',
    rows,
    totals: mfRow('Total', 0, 0),
    wardRows,
    wardTotals: wardRows ? mfRow('Total', 0, 0) : undefined
  }
}

function workRow(
  typeOfWork: string,
  estimateAmt: number,
  overrides: Partial<MonitoringFormatWorkRow> = {}
): MonitoringFormatWorkRow {
  return {
    slNo: '1',
    circle: '',
    ward: '',
    workName: 'Some work',
    estimateAmt,
    typeOfWork,
    sourceOfSanction: '',
    sanctionDate: '',
    agreementDate: '',
    targetDate: '',
    agencyDetails: '',
    status: 'Completed',
    statusKey: 'completed',
    winCode: '',
    ...overrides
  }
}

describe('findMonitoringFormatMismatches', () => {
  it('finds nothing wrong when the Abstract totals match the list of works', () => {
    const s = summary([mfRow('CC Roads', 2, 30)])
    const works = [workRow('CC Roads', 10), workRow('CC Roads', 20)]
    expect(findMonitoringFormatMismatches(s, works)).toEqual([])
  })

  it('flags a count mismatch between the Abstract and the list of works', () => {
    const s = summary([mfRow('CC Roads', 3, 30)])
    const works = [workRow('CC Roads', 10), workRow('CC Roads', 20)]
    const violations = findMonitoringFormatMismatches(s, works)
    expect(violations).toHaveLength(1)
    expect(violations[0]).toMatchObject({ type: 'mf-abstract-mismatch', key: 'CC Roads', rowIndices: [] })
    expect(violations[0].message).toContain('3 works in the Abstract vs 2 in the list of works')
  })

  it('flags an amount mismatch even when the count matches', () => {
    const s = summary([mfRow('SWD', 2, 100)])
    const works = [workRow('SWD', 10), workRow('SWD', 20)]
    const violations = findMonitoringFormatMismatches(s, works)
    expect(violations).toHaveLength(1)
    expect(violations[0].message).toContain('Rs 100 in the Abstract vs Rs 30 in the list of works')
  })

  it('matches item types case/whitespace-insensitively', () => {
    const s = summary([mfRow('CC  Roads', 1, 10)])
    const works = [workRow('cc roads', 10)]
    expect(findMonitoringFormatMismatches(s, works)).toEqual([])
  })

  it('never flags the sheet\'s own "Total" row', () => {
    const s = summary([mfRow('Total', 5, 500)])
    expect(findMonitoringFormatMismatches(s, [])).toEqual([])
  })

  it('skips the ward-wise tally when the sheet has no ward-wise block', () => {
    const s = summary([mfRow('CC Roads', 1, 10)])
    const works = [workRow('CC Roads', 10)]
    expect(findMonitoringFormatMismatches(s, works)).toEqual([])
  })

  it('also tallies the ward-wise pivot against the list of works, when present', () => {
    const s = summary([mfRow('CC Roads', 2, 30)], [mfRow('Bachupally', 2, 30)])
    const works = [
      workRow('CC Roads', 10, { ward: 'Bachupally' }),
      workRow('CC Roads', 20, { ward: 'Bachupally' })
    ]
    expect(findMonitoringFormatMismatches(s, works)).toEqual([])
  })

  it('flags a ward-wise count mismatch even when the item-type tally is clean', () => {
    const s = summary([mfRow('CC Roads', 2, 30)], [mfRow('Bachupally', 3, 30)])
    const works = [
      workRow('CC Roads', 10, { ward: 'Bachupally' }),
      workRow('CC Roads', 20, { ward: 'Bachupally' })
    ]
    const violations = findMonitoringFormatMismatches(s, works)
    expect(violations).toHaveLength(1)
    expect(violations[0].key).toBe('Bachupally')
    expect(violations[0].message).toContain('3 works in the Abstract vs 2 in the list of works')
  })
})

describe('findMonitoringFormatWorksIdentityViolations', () => {
  it('finds nothing wrong with distinct works', () => {
    const works = [
      workRow('CC Roads', 10, { winCode: 'D058-1', workName: 'Work A' }),
      workRow('CC Roads', 20, { winCode: 'D058-2', workName: 'Work B' })
    ]
    expect(findMonitoringFormatWorksIdentityViolations(works)).toEqual([])
  })

  it('flags one Wincode used for two different works', () => {
    const works = [
      workRow('CC Roads', 10, { winCode: 'D058-1', workName: 'Work A' }),
      workRow('CC Roads', 20, { winCode: 'D058-1', workName: 'Work B' })
    ]
    const violations = findMonitoringFormatWorksIdentityViolations(works)
    expect(violations).toHaveLength(1)
    expect(violations[0].type).toBe('duplicate-wincode')
  })

  it('flags one work name entered under two different Wincodes', () => {
    const works = [
      workRow('CC Roads', 10, { winCode: 'D058-1', workName: 'Work A' }),
      workRow('CC Roads', 20, { winCode: 'D058-2', workName: 'Work A' })
    ]
    const violations = findMonitoringFormatWorksIdentityViolations(works)
    expect(violations).toHaveLength(1)
    expect(violations[0].type).toBe('duplicate-name')
  })

  it('does not flag a work correctly re-entered under the same Wincode with a recall tag', () => {
    const works = [
      workRow('CC Roads', 10, { winCode: 'D058-1', workName: 'Work A' }),
      workRow('CC Roads', 20, { winCode: 'D058-1', workName: 'Work A (2nd Recall)' })
    ]
    expect(findMonitoringFormatWorksIdentityViolations(works)).toEqual([])
  })

  it('ignores blank Wincodes/names', () => {
    const works = [
      workRow('CC Roads', 10, { winCode: '', workName: '' }),
      workRow('CC Roads', 20, { winCode: '', workName: '' })
    ]
    expect(findMonitoringFormatWorksIdentityViolations(works)).toEqual([])
  })
})

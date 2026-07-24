import { describe, expect, it } from 'vitest'
import { findCircleSheet, mergeMonitoringRows } from '../core/monitoringImport'
import type { PlaceholderMatch } from '../core/createDocument'
import type { ExcelTable } from '../core/types'
import type { SheetGrid } from '../core/sheet'

function sheet(sheetName: string): SheetGrid {
  return { id: 's1', name: 'Monitoring', path: '', sheetName, grid: [], startRow: 0 }
}

describe('findCircleSheet', () => {
  it('picks the sheet with an exact (normalized) name match', () => {
    const sheets = [sheet('Alwal Circle-12'), sheet('Gajularamaram Circle-57'), sheet('Kukatpally Circle-3')]
    const found = findCircleSheet(sheets, 'Gajularamaram Circle-57')
    expect(found?.sheetName).toBe('Gajularamaram Circle-57')
  })

  it('matches case/whitespace-insensitively', () => {
    const sheets = [sheet('  gajularamaram circle-57 ')]
    const found = findCircleSheet(sheets, 'Gajularamaram Circle-57')
    expect(found?.sheetName).toBe('  gajularamaram circle-57 ')
  })

  it('falls back to a fuzzy contains-match when the sheet name abbreviates the circle', () => {
    const sheets = [sheet('Alwal-12'), sheet('Gajularamaram-57')]
    const found = findCircleSheet(sheets, 'Gajularamaram Circle-57')
    expect(found?.sheetName).toBe('Gajularamaram-57')
  })

  it('returns null when nothing matches', () => {
    const sheets = [sheet('Alwal Circle-12')]
    expect(findCircleSheet(sheets, 'Gajularamaram Circle-57')).toBeNull()
  })
})

function table(headers: string[], rows: Record<string, string>[]): ExcelTable {
  return { id: 't1', name: 'Works database', path: '', headers, rows }
}

describe('mergeMonitoringRows', () => {
  const headers = ['Wincode', 'Name of the work', 'Amount of estimate', 'Contract Amount']
  const mapping: PlaceholderMatch[] = [
    { label: 'Wincode', column: 'Win Code', score: 1 },
    { label: 'Name of the work', column: 'Work Name', score: 1 },
    { label: 'Amount of estimate', column: 'Estimate Amt', score: 1 },
    { label: 'Contract Amount', column: null, score: 0 }
  ]

  it('fills only blank cells on a matched row (by Wincode), never overwriting existing data', () => {
    const t = table(headers, [
      { Wincode: 'WC-1', 'Name of the work': 'Road A', 'Amount of estimate': '', 'Contract Amount': '5,00,000' }
    ])
    const monitoringRows = [{ 'Win Code': 'WC-1', 'Work Name': 'Road A (renamed)', 'Estimate Amt': '4,50,000' }]
    const result = mergeMonitoringRows(t, monitoringRows, mapping)

    expect(result.matchedCount).toBe(1)
    expect(result.addedCount).toBe(0)
    expect(result.filledCount).toBe(1)
    expect(result.table.rows[0]['Amount of estimate']).toBe('4,50,000')
    // Already-filled cells stay exactly as they were.
    expect(result.table.rows[0]['Name of the work']).toBe('Road A')
    expect(result.table.rows[0]['Contract Amount']).toBe('5,00,000')
  })

  it('falls back to matching by Name of the work when Wincode is blank', () => {
    const t = table(headers, [
      { Wincode: '', 'Name of the work': 'Road B', 'Amount of estimate': '', 'Contract Amount': '' }
    ])
    const monitoringRows = [{ 'Win Code': '', 'Work Name': 'Road B', 'Estimate Amt': '2,00,000' }]
    const result = mergeMonitoringRows(t, monitoringRows, mapping)

    expect(result.matchedCount).toBe(1)
    expect(result.table.rows[0]['Amount of estimate']).toBe('2,00,000')
  })

  it('adds a new row for a monitoring row matching no existing Wincode/Name of the work', () => {
    const t = table(headers, [{ Wincode: 'WC-1', 'Name of the work': 'Road A', 'Amount of estimate': '', 'Contract Amount': '' }])
    const monitoringRows = [{ 'Win Code': 'WC-2', 'Work Name': 'Road C (new)', 'Estimate Amt': '3,00,000' }]
    const result = mergeMonitoringRows(t, monitoringRows, mapping)

    expect(result.matchedCount).toBe(0)
    expect(result.addedCount).toBe(1)
    expect(result.table.rows).toHaveLength(2)
    expect(result.table.rows[1]).toMatchObject({
      Wincode: 'WC-2',
      'Name of the work': 'Road C (new)',
      'Amount of estimate': '3,00,000'
    })
  })

  it('skips a monitoring row with neither a Wincode nor a Name of the work to match or add on', () => {
    const t = table(headers, [{ Wincode: 'WC-1', 'Name of the work': 'Road A', 'Amount of estimate': '', 'Contract Amount': '' }])
    const monitoringRows = [{ 'Win Code': '', 'Work Name': '', 'Estimate Amt': '9,99,999' }]
    const result = mergeMonitoringRows(t, monitoringRows, mapping)

    expect(result.matchedCount).toBe(0)
    expect(result.addedCount).toBe(0)
    expect(result.table.rows).toHaveLength(1)
  })

  it('handles several monitoring rows in one merge, matching, filling, and adding correctly', () => {
    const t = table(headers, [
      { Wincode: 'WC-1', 'Name of the work': 'Road A', 'Amount of estimate': '', 'Contract Amount': '' },
      { Wincode: 'WC-2', 'Name of the work': 'Road B', 'Amount of estimate': '1,00,000', 'Contract Amount': '' }
    ])
    const monitoringRows = [
      { 'Win Code': 'WC-1', 'Work Name': 'Road A', 'Estimate Amt': '4,00,000' },
      { 'Win Code': 'WC-2', 'Work Name': 'Road B', 'Estimate Amt': '9,00,000' }, // should NOT overwrite
      { 'Win Code': 'WC-3', 'Work Name': 'Road D', 'Estimate Amt': '2,00,000' }
    ]
    const result = mergeMonitoringRows(t, monitoringRows, mapping)

    expect(result.matchedCount).toBe(2)
    expect(result.addedCount).toBe(1)
    expect(result.table.rows[0]['Amount of estimate']).toBe('4,00,000')
    expect(result.table.rows[1]['Amount of estimate']).toBe('1,00,000') // untouched
    expect(result.table.rows[2]['Wincode']).toBe('WC-3')
  })
})

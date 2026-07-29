import { describe, expect, it } from 'vitest'
import {
  findCircleSheet,
  mergeMonitoringRows,
  splitAgencyNameAndPhones,
  formatAgencyPhones,
  splitCircleNumberAndName,
  looksLikeReferenceEntry
} from '../core/monitoringImport'
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

  it('prefers matching by circle number ("c57"/"C-58" style sheet names) when a circle number is given', () => {
    const sheets = [sheet('c12'), sheet('C-57'), sheet('c 3')]
    expect(findCircleSheet(sheets, 'Gajularamaram Circle-57', '57')?.sheetName).toBe('C-57')
    expect(findCircleSheet(sheets, 'Alwal Circle-12', '12')?.sheetName).toBe('c12')
  })

  it('matches a sheet named with the bare circle number and nothing else', () => {
    const sheets = [sheet('58'), sheet('12')]
    expect(findCircleSheet(sheets, 'Nizampet Circle-58', '58')?.sheetName).toBe('58')
  })

  it('falls back to name-based matching when the circle number matches no sheet', () => {
    const sheets = [sheet('Gajularamaram Circle-57')]
    expect(findCircleSheet(sheets, 'Gajularamaram Circle-57', '99')?.sheetName).toBe('Gajularamaram Circle-57')
  })

  it('prefers the "list of works" sheet over a same-numbered "MF" summary sheet with no per-work rows', () => {
    // Real shape of this app's own monitoring workbook: each circle number
    // pairs an aggregated dashboard sheet ("MF") with the real per-work data
    // ("list of works") — "MF" must never win, in either array order.
    expect(findCircleSheet([sheet('C57 MF'), sheet('C57 list of works')], 'Gajularamaram', '57')?.sheetName).toBe(
      'C57 list of works'
    )
    expect(findCircleSheet([sheet('C58 list of works'), sheet('C58 MF')], 'Nizampet', '58')?.sheetName).toBe(
      'C58 list of works'
    )
  })
})

describe('splitCircleNumberAndName', () => {
  it('splits "<digits>-<name>" into circle number and circle', () => {
    expect(splitCircleNumberAndName('58-Nizampet')).toEqual({ cno: '58', circle: 'Nizampet' })
  })

  it('handles spaces around the hyphen and a multi-word circle name', () => {
    expect(splitCircleNumberAndName('51 - Allwyn Colony')).toEqual({ cno: '51', circle: 'Allwyn Colony' })
  })

  it('splits the number when it comes first with a space ("57 Gajularamaram")', () => {
    expect(splitCircleNumberAndName('57 Gajularamaram')).toEqual({ cno: '57', circle: 'Gajularamaram' })
  })

  it('splits the number when it comes last, hyphen or space ("Gajularamaram-57", "Gajularamaram 57")', () => {
    expect(splitCircleNumberAndName('Gajularamaram-57')).toEqual({ cno: '57', circle: 'Gajularamaram' })
    expect(splitCircleNumberAndName('Gajularamaram 57')).toEqual({ cno: '57', circle: 'Gajularamaram' })
  })

  it('drops a stray "Circle" word so the bare name lands in Circle', () => {
    expect(splitCircleNumberAndName('Gajularamaram Circle-57')).toEqual({ cno: '57', circle: 'Gajularamaram' })
  })

  it('returns null for a value with no circle number', () => {
    expect(splitCircleNumberAndName('Nizampet')).toBeNull()
    expect(splitCircleNumberAndName('')).toBeNull()
  })
})

describe('splitAgencyNameAndPhones', () => {
  it('splits a single 10-digit phone number out of the agency name', () => {
    const { name, phones } = splitAgencyNameAndPhones('Radha Krishna Contractors 9789879878')
    expect(name).toBe('Radha Krishna Contractors')
    expect(phones).toEqual(['9789879878'])
  })

  it('splits two phone numbers, keeping their order', () => {
    const { name, phones } = splitAgencyNameAndPhones('Radha Krishna Contractors 9789879878 9848012345')
    expect(name).toBe('Radha Krishna Contractors')
    expect(phones).toEqual(['9789879878', '9848012345'])
  })

  it('strips hyphens/dots within a single phone token before counting digits', () => {
    const { phones } = splitAgencyNameAndPhones('Radha Krishna Contractors 978-987-9878')
    expect(phones).toEqual(['9789879878'])
  })

  it('leaves a name with no embedded phone number untouched', () => {
    const { name, phones } = splitAgencyNameAndPhones('Radha Krishna Contractors')
    expect(name).toBe('Radha Krishna Contractors')
    expect(phones).toEqual([])
  })

  it('leaves a non-10-digit number (e.g. a registration number) in the name rather than misidentifying it', () => {
    const { name, phones } = splitAgencyNameAndPhones('Radha Krishna Contractors REG12345')
    expect(name).toBe('Radha Krishna Contractors REG12345')
    expect(phones).toEqual([])
  })

  it('drops a trailing comma left over after removing a phone number', () => {
    const { name } = splitAgencyNameAndPhones('Radha Krishna Contractors, 9789879878')
    expect(name).toBe('Radha Krishna Contractors')
  })
})

describe('formatAgencyPhones', () => {
  it('returns the bare number for exactly one phone', () => {
    expect(formatAgencyPhones(['9789879878'])).toBe('9789879878')
  })

  it('numbers each phone on its own line for more than one', () => {
    expect(formatAgencyPhones(['9789879878', '9848012345'])).toBe('1. 9789879878\n2. 9848012345')
  })

  it('returns an empty string for no phones', () => {
    expect(formatAgencyPhones([])).toBe('')
  })
})

describe('looksLikeReferenceEntry', () => {
  it('flags a circle/ward dropdown token (no Wincode) as a reference entry', () => {
    expect(looksLikeReferenceEntry('54-Chintal', '')).toBe(true)
    expect(looksLikeReferenceEntry('57-Gajularamaram', '')).toBe(true)
    expect(looksLikeReferenceEntry('277-Mahadevapuram', '')).toBe(true)
  })

  it('does not flag a real, descriptive work name', () => {
    expect(looksLikeReferenceEntry('Laying of CC road at Ayyappa Colony, Gajularamaram Circle', '')).toBe(false)
    expect(looksLikeReferenceEntry('2 lane road widening from X to Y in ward 12', '')).toBe(false)
  })

  it('does not flag anything that carries a Wincode (a real work)', () => {
    expect(looksLikeReferenceEntry('58-Nizampet', 'WC-123')).toBe(false)
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

  it('does not add a monitoring sheet\'s dropdown-reference entries (circle/ward lists) as works', () => {
    const t = table(headers, [{ Wincode: 'WC-1', 'Name of the work': 'Road A', 'Amount of estimate': '', 'Contract Amount': '' }])
    // A legend block: circle list mapped onto Name of the work, ward list onto
    // Amount of estimate — no Wincodes. None of these should become works.
    const monitoringRows = [
      { 'Win Code': '', 'Work Name': '54-Chintal', 'Estimate Amt': '277-Mahadevapuram' },
      { 'Win Code': '', 'Work Name': '57-Gajularamaram', 'Estimate Amt': '291-Shapur Nagar' },
      { 'Win Code': '', 'Work Name': '58-Nizampet', 'Estimate Amt': '293-Suraram' }
    ]
    const result = mergeMonitoringRows(t, monitoringRows, mapping)
    expect(result.addedCount).toBe(0)
    expect(result.table.rows).toHaveLength(1)
  })

  it('still adds a real work whose name happens to start with a number (has a Wincode / long name)', () => {
    const t = table(headers, [{ Wincode: 'WC-1', 'Name of the work': 'Road A', 'Amount of estimate': '', 'Contract Amount': '' }])
    const monitoringRows = [
      { 'Win Code': 'WC-9', 'Work Name': '2 lane road widening from X to Y in ward 12', 'Estimate Amt': '5,00,000' }
    ]
    const result = mergeMonitoringRows(t, monitoringRows, mapping)
    expect(result.addedCount).toBe(1)
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

  it('splits an embedded phone number out of the Name of the Agency column into Phone number of the agency', () => {
    const agencyHeaders = [...headers, 'Name of the Agency', 'Phone number of the agency']
    const agencyMapping: PlaceholderMatch[] = [...mapping, { label: 'Name of the Agency', column: 'Agency', score: 1 }]
    const t = table(agencyHeaders, [
      {
        Wincode: 'WC-1',
        'Name of the work': 'Road A',
        'Amount of estimate': '',
        'Contract Amount': '',
        'Name of the Agency': '',
        'Phone number of the agency': ''
      }
    ])
    const monitoringRows = [{ 'Win Code': 'WC-1', 'Work Name': 'Road A', Agency: 'Radha Krishna Contractors 9789879878 9848012345' }]
    const result = mergeMonitoringRows(t, monitoringRows, agencyMapping)

    expect(result.table.rows[0]['Name of the Agency']).toBe('Radha Krishna Contractors')
    expect(result.table.rows[0]['Phone number of the agency']).toBe('1. 9789879878\n2. 9848012345')
  })

  it('does not touch an already-filled Phone number of the agency even when the agency name has its own embedded number', () => {
    const agencyHeaders = [...headers, 'Name of the Agency', 'Phone number of the agency']
    const agencyMapping: PlaceholderMatch[] = [...mapping, { label: 'Name of the Agency', column: 'Agency', score: 1 }]
    const t = table(agencyHeaders, [
      {
        Wincode: 'WC-1',
        'Name of the work': 'Road A',
        'Amount of estimate': '',
        'Contract Amount': '',
        'Name of the Agency': '',
        'Phone number of the agency': '9000000000'
      }
    ])
    const monitoringRows = [{ 'Win Code': 'WC-1', 'Work Name': 'Road A', Agency: 'Radha Krishna Contractors 9789879878' }]
    const result = mergeMonitoringRows(t, monitoringRows, agencyMapping)

    expect(result.table.rows[0]['Name of the Agency']).toBe('Radha Krishna Contractors')
    expect(result.table.rows[0]['Phone number of the agency']).toBe('9000000000')
  })

  it('splits a combined "58-Nizampet" Circle column into Circle and Circle number separately', () => {
    const circleHeaders = [...headers, 'Circle', 'Circle number']
    const circleMapping: PlaceholderMatch[] = [...mapping, { label: 'Circle', column: 'Name of the Circle', score: 1 }]
    const t = table(circleHeaders, [
      {
        Wincode: 'WC-1',
        'Name of the work': 'Road A',
        'Amount of estimate': '',
        'Contract Amount': '',
        Circle: '',
        'Circle number': ''
      }
    ])
    const monitoringRows = [{ 'Win Code': 'WC-1', 'Work Name': 'Road A', 'Name of the Circle': '58-Nizampet' }]
    const result = mergeMonitoringRows(t, monitoringRows, circleMapping)

    expect(result.table.rows[0]['Circle']).toBe('Nizampet')
    expect(result.table.rows[0]['Circle number']).toBe('58')
  })
})

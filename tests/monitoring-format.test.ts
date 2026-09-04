import { describe, expect, it } from 'vitest'
import {
  parseMonitoringFormatSheet,
  findMonitoringFormatSheet,
  extractMonitoringFormatForOffice
} from '../core/monitoringFormat'
import type { SheetGrid } from '../core/sheet'

// Mirrors the real "MF" sheet layout: a "Name of the Cir(c)le:-"/"Date:-"
// label row, three header rows (group / sub-group / No.-Amt.), then one row
// per item type, ending in the sheet's own "Total" row.
function mfSheet(sheetName: string, officeLabel: string, rows: string[][]): SheetGrid {
  const grid = [
    ['Monitoring Format 2026-27'],
    ['Name of the Cir(c)le:-', officeLabel, '', 'Date:-', '23.07.2026'],
    [],
    [
      'Item type',
      'Total Works',
      '',
      'Completed',
      '',
      'upto 25%',
      '',
      'upto 50%',
      '',
      'upto 75%',
      '',
      'Above 75%',
      '',
      'Progress Total',
      '',
      'To be Started',
      '',
      'Tender Process',
      '',
      'Held Up',
      '',
      'Cancelled',
      ''
    ],
    [],
    ['', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt', 'No.', 'Amt'],
    ...rows
  ]
  return { id: 's1', name: 'MF', path: '', sheetName, grid, startRow: 0 }
}

const ccRoadsRow = ['CC ROADS', '95', '2,473.5', '44', '1,107.5', '5', '178', '4', '127', '0', '0', '0', '0', '9', '305', '26', '638', '9', '231', '0', '0', '0', '0']
const swdRow = ['SWD', '10', '500', '5', '250', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '3', '150', '2', '100', '0', '0', '0', '0']
const totalRow = ['Total', '105', '2973.5', '49', '1357.5', '5', '178', '4', '127', '0', '0', '0', '0', '9', '305', '29', '788', '11', '331', '0', '0', '0', '0']

describe('parseMonitoringFormatSheet', () => {
  it('parses item-type rows and the sheet\'s own Total row', () => {
    const sheet = mfSheet('C58 MF', '58-Nizampet', [ccRoadsRow, swdRow, totalRow])
    const summary = parseMonitoringFormatSheet(sheet)

    expect(summary.officeLabel).toBe('58-Nizampet')
    expect(summary.asOfDate).toBe('23.07.2026')
    expect(summary.rows.map((r) => r.itemType)).toEqual(['CC ROADS', 'SWD'])
    expect(summary.rows[0].totalWorks).toEqual({ no: 95, amt: 2473.5 })
    expect(summary.rows[0].completed).toEqual({ no: 44, amt: 1107.5 })
    expect(summary.rows[0].toBeStarted).toEqual({ no: 26, amt: 638 })
    expect(summary.totals).toEqual({
      itemType: 'Total',
      totalWorks: { no: 105, amt: 2973.5 },
      completed: { no: 49, amt: 1357.5 },
      upto25: { no: 5, amt: 178 },
      upto50: { no: 4, amt: 127 },
      upto75: { no: 0, amt: 0 },
      above75: { no: 0, amt: 0 },
      progressTotal: { no: 9, amt: 305 },
      toBeStarted: { no: 29, amt: 788 },
      tenderProcess: { no: 11, amt: 331 },
      heldUp: { no: 0, amt: 0 },
      cancelled: { no: 0, amt: 0 }
    })
  })

  it('sums the item-type rows itself when the sheet has no Total row', () => {
    const sheet = mfSheet('C58 MF', '58-Nizampet', [ccRoadsRow, swdRow])
    const summary = parseMonitoringFormatSheet(sheet)
    expect(summary.totals.totalWorks).toEqual({ no: 105, amt: 2973.5 })
    expect(summary.totals.toBeStarted).toEqual({ no: 29, amt: 788 })
  })

  it('throws when the sheet has no "Item type" header row', () => {
    const sheet: SheetGrid = { id: 's1', name: 'MF', path: '', sheetName: 'Not MF', grid: [['foo', 'bar']], startRow: 0 }
    expect(() => parseMonitoringFormatSheet(sheet)).toThrow(/doesn't look like a Monitoring Format sheet/)
  })
})

describe('findMonitoringFormatSheet / extractMonitoringFormatForOffice', () => {
  const sheets = [
    mfSheet('QBZ MF', 'Quthbullapur Zone', [ccRoadsRow, totalRow]),
    mfSheet('C58 MF', '58-Nizampet', [ccRoadsRow, totalRow]),
    mfSheet('C57 MF', '57-Gajularamaram', [ccRoadsRow, totalRow]),
    // A "list of works" sheet matching /\bmf\b/i by coincidence must never be
    // picked over the real per-circle "MF" summary sheet.
    { id: 's2', name: 'MF', path: '', sheetName: 'C58 MF list of works', grid: [['Item type']], startRow: 0 }
  ]

  it('picks the sheet matching a circle by number', () => {
    const found = findMonitoringFormatSheet(sheets, { circle: 'Nizampet', circleNumber: '58' })
    expect(found?.sheetName).toBe('C58 MF')
  })

  it('picks the sheet matching a circle by name when the number is missing', () => {
    const found = findMonitoringFormatSheet(sheets, { circle: 'Gajularamaram' })
    expect(found?.sheetName).toBe('C57 MF')
  })

  it('picks the lone zone rollup sheet for a zone-only office', () => {
    const found = findMonitoringFormatSheet(sheets, { zone: 'Quthbullapur' })
    expect(found?.sheetName).toBe('QBZ MF')
  })

  it('returns null when no sheet matches the office', () => {
    expect(findMonitoringFormatSheet(sheets, { circle: 'Somewhere Else', circleNumber: '99' })).toBeNull()
  })

  it('extractMonitoringFormatForOffice parses the matched sheet', () => {
    const summary = extractMonitoringFormatForOffice(sheets, { circle: 'Nizampet', circleNumber: '58' })
    expect(summary.officeLabel).toBe('58-Nizampet')
    expect(summary.sheetName).toBe('C58 MF')
  })

  it('extractMonitoringFormatForOffice throws a helpful error when nothing matches', () => {
    expect(() => extractMonitoringFormatForOffice(sheets, { circle: 'Somewhere Else', circleNumber: '99' })).toThrow(
      /Somewhere Else/
    )
  })
})

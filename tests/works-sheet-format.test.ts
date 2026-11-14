import { describe, it, expect } from 'vitest'
import { aliasMapping, cleanSheetRows, circleOfWork } from '../core/worksSheetFormat'
import { WORKS_COLUMNS } from '../src/worksSchema'
import { updateWorksListFromEvaluations } from '../core/worksTenderUpdate'
import { saveWorkspaceInputs, sanctionFinancialYear } from '../core/worksSync'
import { circleFromWorkName } from '../src/zoneCircleCheck'

const SE_HEADERS = [
  'Sl. No.',
  'Cirlce no. and Name',
  'Name  of the Work',
  'Estimate  Amount',
  'E C V',
  'Nit no and Date',
  'Contract Value',
  '% Tender Percentage',
  'Name of the Contractor',
  'Agency \n Phone No.',
  'Techincal Sanction No.',
  'Techincal Sanction Date',
  'Admin\n Sanction dt',
  'AS Given by',
  'Eoffice File No.',
  'POC in months',
  'wincode',
  'Sl. No./ LOA no.',
  'LOA Date',
  'Agreement no.',
  'Agreement Dt and no.',
  'FY '
]

const colFor = (headers: string[], label: string) =>
  aliasMapping(WORKS_COLUMNS, headers).find((m) => m.label === label)?.column

describe('aliasMapping', () => {
  it('reads the SE office sheet by name, typos and all', () => {
    expect(colFor(SE_HEADERS, 'Circle')).toBe('Cirlce no. and Name')
    expect(colFor(SE_HEADERS, 'ECV')).toBe('E C V')
    expect(colFor(SE_HEADERS, 'Technical Sanc No')).toBe('Techincal Sanction No.')
    expect(colFor(SE_HEADERS, 'TS date')).toBe('Techincal Sanction Date')
    expect(colFor(SE_HEADERS, 'Admin Sanction Date')).toBe('Admin\n Sanction dt')
    expect(colFor(SE_HEADERS, 'Sanction By')).toBe('AS Given by')
    expect(colFor(SE_HEADERS, 'Completion Period')).toBe('POC in months')
    expect(colFor(SE_HEADERS, 'Intimation Date')).toBe('LOA Date')
    expect(colFor(SE_HEADERS, 'Agmt Date')).toBe('Agreement Dt and no.')
    expect(colFor(SE_HEADERS, 'Name of the Agency')).toBe('Name of the Contractor')
  })

  it('keeps a separate Circle number column apart from the Circle column', () => {
    const h = ['Wincode', 'Name of the work', 'Circle', 'Circle number']
    expect(colFor(h, 'Circle')).toBe('Circle')
    expect(colFor(h, 'Circle number')).toBe('Circle number')
  })

  it('claims no column twice and leaves unknown headers unmapped', () => {
    const mapped = aliasMapping(WORKS_COLUMNS, ['Name of the work', 'Remarks']).filter((m) => m.column)
    expect(mapped.map((m) => m.column)).toEqual(['Name of the work'])
  })
})

describe('cleanSheetRows', () => {
  const headers = ['Sl. No.', 'Name of the Work', 'Contract Value']
  it('drops financial-year banner rows and rows with no work', () => {
    const rows = [
      { 'Sl. No.': '2025-26', 'Name of the Work': '', 'Contract Value': '' },
      { 'Sl. No.': '1', 'Name of the Work': 'Road', 'Contract Value': 'Rs.61,67,723.45' },
      { 'Sl. No.': '2', 'Name of the Work': '', 'Contract Value': '' }
    ]
    expect(cleanSheetRows(headers, rows)).toEqual([
      { 'Sl. No.': '1', 'Name of the Work': 'Road', 'Contract Value': '6167723.45' }
    ])
  })

  it('keeps every row when the sheet has no recognisable work column', () => {
    const rows = [{ A: '1' }, { A: '2' }]
    expect(cleanSheetRows(['A'], rows)).toHaveLength(2)
  })
})

describe('circleOfWork', () => {
  const names = { '60': 'Medchal', '54': 'Chintal' }
  it('never alters a filled cell', () => {
    expect(circleOfWork('55-Jeedimetla', 'x Circle 60', names)).toEqual({ circle: '55-Jeedimetla', cno: '55' })
  })
  it('reads a blank circle from the work name', () => {
    expect(circleOfWork('', 'Culvert at Pudur, Medchal Circle 60, QBZ', names)).toEqual({
      circle: '60-Medchal',
      cno: '60'
    })
    expect(circleOfWork('', 'lane marking in chintal circle-54, Zone', names)).toEqual({
      circle: '54-Chintal',
      cno: '54'
    })
  })
  it('leaves it blank when the name carries no circle', () => {
    expect(circleOfWork('', 'Procurement of computers', names)).toEqual({ circle: '', cno: '' })
  })
})

describe('L-1 update fills a blank Circle from the work name', () => {
  const mk = (circle: string) => ({
    id: 't', name: 'w', path: '', headers: ['Circle', 'Name of the work'],
    rows: [{ Circle: circle, 'Name of the work': 'Laying CC road in Nizampet circle, Quthbullapur zone' }]
  })
  const ev = (nameOfWork: string) => ({ nameOfWork }) as never
  const run = (circle: string, name: string) => {
    const t = mk(circle)
    t.rows[0]['Name of the work'] = name
    return updateWorksListFromEvaluations(t, [ev(name)], undefined, undefined, (n) => circleFromWorkName(n, 'CMC')).table.rows[0].Circle
  }
  it('fills the circle the name designates, as the sheet writes it', () => {
    expect(run('', 'Laying CC road in Nizampet circle, Quthbullapur zone')).toMatch(/^\d+-Nizampet$/)
  })
  it('fills Zonal Office for a work named for the zone or the zonal office', () => {
    expect(run('', 'Procurement of computers for Quthbullapur Zone')).toBe('Zonal Office')
    expect(run('', 'Repairs to the Zonal Office building')).toBe('Zonal Office')
  })
  it('leaves it blank when the name says neither', () => {
    expect(run('', 'Procurement of computers')).toBe('')
  })
  it('never rewrites a Circle that is already filled', () => {
    expect(run('55-Jeedimetla', 'Laying CC road in Nizampet circle, Quthbullapur zone')).toBe('55-Jeedimetla')
  })
})

describe('workspace inputs are saved into the sheet’s own columns', () => {
  const table = {
    id: 't', name: 'w', path: '',
    headers: ['Name of the work', 'Sl. No./ LOA no.', 'Completion Period', 'Agreement no.', 'Agreement Dt and no.', 'Financial Year'],
    rows: [{ 'Name of the work': 'Road', 'Sl. No./ LOA no.': '', 'Completion Period': '', 'Agreement no.': '', 'Agreement Dt and no.': '09-03-2026', 'Financial Year': '' }]
  }
  it('fills LOA no, period, agreement no and FY, leaving other columns alone', () => {
    const out = saveWorkspaceInputs(table, 0, { loaNo: '21/SE/QBZ/CMC/2026-27', completionPeriod: '3', agreementNo: '12', financialYear: '2026-27' })
    expect(out.rows[0]).toMatchObject({
      'Sl. No./ LOA no.': '21/SE/QBZ/CMC/2026-27', 'Completion Period': '3', 'Agreement no.': '12', 'Financial Year': '2026-27',
      'Agreement Dt and no.': '09-03-2026'
    })
    expect(out.headers).toEqual(table.headers)
  })
  it('fills the agreement date into the sheet\'s agreement date column', () => {
    const out = saveWorkspaceInputs(table, 0, { agreementDate: '04.04.2026' })
    expect(out.rows[0]['Agreement Dt and no.']).toBe('04.04.2026')
  })
  it('never replaces a financial year that is already there when asked not to', () => {
    const filled = { ...table, rows: [{ ...table.rows[0], 'Financial Year': '2025-2026' }] }
    expect(saveWorkspaceInputs(filled, 0, { financialYear: '2026-27' }, ['financialYear'])).toBe(filled)
  })
  it('adds no column when the database has none for the input', () => {
    const plain = { ...table, headers: ['Name of the work'], rows: [{ 'Name of the work': 'Road' }] }
    expect(saveWorkspaceInputs(plain, 0, { loaNo: '5', agreementNo: '2' })).toBe(plain)
  })
})

describe('sanctionFinancialYear', () => {
  it('is the year of the administrative sanction date, in either date style', () => {
    expect(sanctionFinancialYear('03-02-2026')).toBe('2025-26') // Feb 2026 is in FY 2025-26
    expect(sanctionFinancialYear('15.02.2026')).toBe('2025-26')
    expect(sanctionFinancialYear('10.07.2026')).toBe('2026-27')
    expect(sanctionFinancialYear('31-03-2026')).toBe('2025-26')
    expect(sanctionFinancialYear('01-04-2026')).toBe('2026-27')
  })
  it('is blank when there is no readable administrative sanction date', () => {
    expect(sanctionFinancialYear('-')).toBe('')
    expect(sanctionFinancialYear('HACA')).toBe('')
    expect(sanctionFinancialYear(undefined)).toBe('')
  })
})

describe('L-1 update fills a blank Zone from the work name', () => {
  const run = (zone: string, name: string) => {
    const t = { id: 't', name: 'w', path: '', headers: ['Zone', 'Circle', 'Name of the work'], rows: [{ Zone: zone, Circle: '', 'Name of the work': name }] }
    return updateWorksListFromEvaluations(t, [{ nameOfWork: name } as never], undefined, undefined, (n) => circleFromWorkName(n, 'CMC')).table.rows[0]
  }
  it('fills the zone with the circle, and for a zone-level work', () => {
    expect(run('', 'Laying CC road in Nizampet circle, Quthbullapur zone').Zone).toBe('Quthbullapur')
    expect(run('', 'Procurement of computers for Quthbullapur Zone').Zone).toBe('Quthbullapur')
  })
  it('never rewrites a filled zone, and leaves it blank when the name has none', () => {
    expect(run('Kukatpally', 'Laying CC road in Nizampet circle').Zone).toBe('Kukatpally')
    expect(run('', 'Procurement of computers').Zone).toBe('')
  })
})

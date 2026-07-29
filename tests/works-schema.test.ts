import { describe, expect, it } from 'vitest'
import {
  applyWorksSchema,
  applyWorksSchemaWithMapping,
  migrateEcvContractToRupees,
  repairInflatedRupees,
  WORKS_COLUMNS
} from '../src/worksSchema'
import type { PlaceholderMatch } from '../core/createDocument'
import type { ExcelTable } from '../core/types'

describe('applyWorksSchemaWithMapping', () => {
  it('pulls a standard column from whichever imported header the mapping resolved it to', () => {
    const mapping: PlaceholderMatch[] = [
      { label: 'Name of the work', column: 'Work Name', score: 0.9 },
      { label: 'Amount of estimate', column: 'Estimate Amount', score: 0.8 }
    ]
    const headers = ['Work Name', 'Estimate Amount']
    const rows = [{ 'Work Name': 'Road repair', 'Estimate Amount': '5,00,000' }]
    const result = applyWorksSchemaWithMapping(headers, rows, mapping, { id: 't1', name: 'Works', path: '' })

    expect(result.headers).toEqual(WORKS_COLUMNS)
    expect(result.rows[0]['Name of the work']).toBe('Road repair')
    expect(result.rows[0]['Amount of estimate']).toBe('5,00,000')
  })

  it('falls back to an exact name match for a column the mapping left unresolved', () => {
    const mapping: PlaceholderMatch[] = [{ label: 'Name of the work', column: null, score: 0 }]
    const headers = ['Name of the work']
    const rows = [{ 'Name of the work': 'Bridge work' }]
    const result = applyWorksSchemaWithMapping(headers, rows, mapping, { id: 't1', name: 'Works', path: '' })
    expect(result.rows[0]['Name of the work']).toBe('Bridge work')
  })

  it('leaves a standard column blank when neither the mapping nor an exact name match resolves it', () => {
    const mapping: PlaceholderMatch[] = [{ label: 'Zone', column: null, score: 0 }]
    const headers = ['Some Other Column']
    const rows = [{ 'Some Other Column': 'x' }]
    const result = applyWorksSchemaWithMapping(headers, rows, mapping, { id: 't1', name: 'Works', path: '' })
    expect(result.rows[0]['Zone']).toBe('')
  })

  it('guarantees at least one (blank) row when given none', () => {
    const result = applyWorksSchemaWithMapping([], [], [], { id: 't1', name: 'Works', path: '' })
    expect(result.rows).toHaveLength(1)
    expect(result.rows[0]['Name of the work']).toBe('')
  })

  it('appends an imported column the mapping did not claim for any standard column, keeping its own name', () => {
    const mapping: PlaceholderMatch[] = [
      { label: 'Name of the work', column: 'Name of the work', score: 1 },
      { label: 'Zone', column: null, score: 0 }
    ]
    const headers = ['Name of the work', 'Eoffice', 'Download start time']
    const rows = [{ 'Name of the work': 'Road A', Eoffice: 'EO-123', 'Download start time': '10:00' }]
    const result = applyWorksSchemaWithMapping(headers, rows, mapping, { id: 't1', name: 'Works', path: '' })

    expect(result.headers).toEqual([...WORKS_COLUMNS, 'Eoffice', 'Download start time'])
    expect(result.rows[0]['Eoffice']).toBe('EO-123')
    expect(result.rows[0]['Download start time']).toBe('10:00')
  })

  it('does not duplicate an imported column that was claimed by a standard column', () => {
    const mapping: PlaceholderMatch[] = [{ label: 'Amount of estimate', column: 'Estimate Amount', score: 0.9 }]
    const headers = ['Estimate Amount']
    const rows = [{ 'Estimate Amount': '1,00,000' }]
    const result = applyWorksSchemaWithMapping(headers, rows, mapping, { id: 't1', name: 'Works', path: '' })

    expect(result.headers).toEqual(WORKS_COLUMNS)
    expect(result.headers).not.toContain('Estimate Amount')
  })
})

function table(headers: string[], rows: Record<string, string>[]): ExcelTable {
  return { id: 't1', name: 'Works database', path: '', headers, rows }
}

describe('applyWorksSchema', () => {
  it('guarantees every standard column exists', () => {
    const t = table(['Name of the work'], [{ 'Name of the work': 'Road A' }])
    const result = applyWorksSchema(t)
    expect(result.headers).toEqual(WORKS_COLUMNS)
  })

  it('preserves an extra (non-standard) column already on the table instead of dropping it', () => {
    const t = table(
      [...WORKS_COLUMNS, 'Eoffice'],
      [{ ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])), Eoffice: 'EO-123' }]
    )
    const result = applyWorksSchema(t)
    expect(result.headers).toEqual([...WORKS_COLUMNS, 'Eoffice'])
    expect(result.rows[0]['Eoffice']).toBe('EO-123')
  })

  it('carries a table saved under the old "Estimate Amount ECV" column name into the new "ECV" column', () => {
    const headers = WORKS_COLUMNS.map((h) => (h === 'ECV' ? 'Estimate Amount ECV' : h))
    const t = table(headers, [{ ...Object.fromEntries(headers.map((h) => [h, ''])), 'Estimate Amount ECV': '45' }])
    const result = applyWorksSchema(t)
    expect(result.rows[0]['ECV']).toBe('45')
  })

  it('does not overwrite a real "ECV" value with a stale legacy column of the same row', () => {
    const t = table(WORKS_COLUMNS, [{ ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])), ECV: '50' }])
    const result = applyWorksSchema(t)
    expect(result.rows[0]['ECV']).toBe('50')
  })
})

describe('migrateEcvContractToRupees', () => {
  it('multiplies ECV and Contract Amount (Lakhs) by 100000, leaving other columns alone', () => {
    const t = table(WORKS_COLUMNS, [
      {
        ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])),
        'Amount of estimate': '45',
        ECV: '15.93493',
        'Contract Amount': '14.16456'
      }
    ])
    const out = migrateEcvContractToRupees(t)
    expect(out.rows[0]['ECV']).toBe('1593493')
    expect(out.rows[0]['Contract Amount']).toBe('1416456')
    // Amount of estimate stays in Lakhs, untouched.
    expect(out.rows[0]['Amount of estimate']).toBe('45')
  })

  it('leaves blank and non-numeric ECV/Contract cells untouched', () => {
    const t = table(WORKS_COLUMNS, [
      { ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])), ECV: '', 'Contract Amount': 'N/A' }
    ])
    const out = migrateEcvContractToRupees(t)
    expect(out.rows[0]['ECV']).toBe('')
    expect(out.rows[0]['Contract Amount']).toBe('N/A')
  })

  it('does NOT re-inflate an already-rupees value (guard against a repeated migration)', () => {
    const t = table(WORKS_COLUMNS, [
      { ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])), ECV: '2571292', 'Contract Amount': '1735622' }
    ])
    const out = migrateEcvContractToRupees(t)
    expect(out.rows[0]['ECV']).toBe('2571292')
    expect(out.rows[0]['Contract Amount']).toBe('1735622')
  })
})

describe('repairInflatedRupees', () => {
  it('divides an over-inflated ECV/Contract Amount back to a plausible rupee figure', () => {
    const t = table(WORKS_COLUMNS, [
      { ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])), ECV: '2.571292e+21', 'Contract Amount': '1735622000000000000000' }
    ])
    const out = repairInflatedRupees(t)
    expect(out.rows[0]['ECV']).toBe('2571292')
    expect(out.rows[0]['Contract Amount']).toBe('1735622')
  })

  it('leaves already-plausible rupee values untouched (idempotent)', () => {
    const t = table(WORKS_COLUMNS, [
      { ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])), ECV: '880637', 'Contract Amount': '725205' }
    ])
    const out = repairInflatedRupees(t)
    expect(out.rows[0]['ECV']).toBe('880637')
    expect(out.rows[0]['Contract Amount']).toBe('725205')
  })
})

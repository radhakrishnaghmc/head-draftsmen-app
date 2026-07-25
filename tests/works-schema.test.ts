import { describe, expect, it } from 'vitest'
import { applyWorksSchema, applyWorksSchemaWithMapping, WORKS_COLUMNS } from '../src/worksSchema'
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

import { describe, expect, it } from 'vitest'
import { applyWorksSchemaWithMapping, WORKS_COLUMNS } from '../src/worksSchema'
import type { PlaceholderMatch } from '../core/createDocument'

describe('applyWorksSchemaWithMapping', () => {
  it('pulls a standard column from whichever imported header the mapping resolved it to', () => {
    const mapping: PlaceholderMatch[] = [
      { label: 'Name of the work', column: 'Work Name', score: 0.9 },
      { label: 'Amount of estimate', column: 'Estimate Amount', score: 0.8 }
    ]
    const rows = [{ 'Work Name': 'Road repair', 'Estimate Amount': '5,00,000' }]
    const result = applyWorksSchemaWithMapping(rows, mapping, { id: 't1', name: 'Works', path: '' })

    expect(result.headers).toEqual(WORKS_COLUMNS)
    expect(result.rows[0]['Name of the work']).toBe('Road repair')
    expect(result.rows[0]['Amount of estimate']).toBe('5,00,000')
  })

  it('falls back to an exact name match for a column the mapping left unresolved', () => {
    const mapping: PlaceholderMatch[] = [{ label: 'Name of the work', column: null, score: 0 }]
    const rows = [{ 'Name of the work': 'Bridge work' }]
    const result = applyWorksSchemaWithMapping(rows, mapping, { id: 't1', name: 'Works', path: '' })
    expect(result.rows[0]['Name of the work']).toBe('Bridge work')
  })

  it('leaves a standard column blank when neither the mapping nor an exact name match resolves it', () => {
    const mapping: PlaceholderMatch[] = [{ label: 'Zone', column: null, score: 0 }]
    const rows = [{ 'Some Other Column': 'x' }]
    const result = applyWorksSchemaWithMapping(rows, mapping, { id: 't1', name: 'Works', path: '' })
    expect(result.rows[0]['Zone']).toBe('')
  })

  it('guarantees at least one (blank) row when given none', () => {
    const result = applyWorksSchemaWithMapping([], [], { id: 't1', name: 'Works', path: '' })
    expect(result.rows).toHaveLength(1)
    expect(result.rows[0]['Name of the work']).toBe('')
  })
})

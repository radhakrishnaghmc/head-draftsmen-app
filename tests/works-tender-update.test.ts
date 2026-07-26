import { describe, expect, it } from 'vitest'
import { updateWorksListFromEvaluations, applyAgencyAddresses } from '../core/worksTenderUpdate'
import type { TenderEvaluation } from '../core/tenderEvaluationPdf'
import type { ExcelTable } from '../core/types'
import { WORKS_COLUMNS } from '../src/worksSchema'

function blankRow(overrides: Record<string, string> = {}): Record<string, string> {
  return { ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])), ...overrides }
}
function table(rows: Record<string, string>[]): ExcelTable {
  return { id: 't1', name: 'Works database', path: '', headers: WORKS_COLUMNS, rows }
}

const EV: TenderEvaluation = {
  nameOfWork: 'Junction Improvement in Aleap Circle',
  tenderId: '717574',
  noticeNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27',
  noticeDate: '15.07.2026',
  ecvRupees: 1593493,
  l1AgencyName: 'M V S CONSTRUCTIONS',
  tenderPercentage: 11.11,
  contractRupees: 1416455.93
}

describe('updateWorksListFromEvaluations', () => {
  it('fills all six fields on the row whose work name matches (case/space-insensitive)', () => {
    const t = table([blankRow({ 'Name of the work': '  junction   improvement in aleap circle ' })])
    const { table: out, matchedCount, unmatched } = updateWorksListFromEvaluations(t, [EV])
    expect(matchedCount).toBe(1)
    expect(unmatched).toEqual([])
    const r = out.rows[0]
    expect(r['Tender ID']).toBe('717574')
    expect(r['Tender Notice No']).toBe('12/DB/EE/Nizampet Circle-58/CMC/2026-27')
    expect(r['Tender notice Date']).toBe('15.07.2026')
    expect(r['ECV']).toBe('1593493') // stored in rupees
    expect(r['Name of the Agency']).toBe('M V S CONSTRUCTIONS')
    expect(r['Tender Percentage']).toBe('11.11')
    expect(r['Contract Amount']).toBe('1416456') // rupees, rounded to the nearest rupee
  })

  it('overwrites existing values (the PDF is the authoritative post-award record)', () => {
    const t = table([
      blankRow({ 'Name of the work': 'Junction Improvement in Aleap Circle', 'Tender ID': 'old', ECV: '99' })
    ])
    const { table: out } = updateWorksListFromEvaluations(t, [EV])
    expect(out.rows[0]['Tender ID']).toBe('717574')
    expect(out.rows[0]['ECV']).toBe('1593493')
  })

  it('only writes fields the PDF carried, leaving the rest of the row alone', () => {
    const t = table([blankRow({ 'Name of the work': 'Road work X', Zone: 'Cyberabad', 'EMD 1%': '5' })])
    const partial: TenderEvaluation = { nameOfWork: 'Road work X', tenderId: 'T-9' }
    const { table: out } = updateWorksListFromEvaluations(t, [partial])
    expect(out.rows[0]['Tender ID']).toBe('T-9')
    expect(out.rows[0]['Zone']).toBe('Cyberabad')
    expect(out.rows[0]['EMD 1%']).toBe('5')
    expect(out.rows[0]['Name of the Agency']).toBe('')
  })

  it('reports a PDF whose work name matches no row as unmatched, touching nothing', () => {
    const t = table([blankRow({ 'Name of the work': 'Totally different work' })])
    const { table: out, matchedCount, unmatched } = updateWorksListFromEvaluations(t, [EV])
    expect(matchedCount).toBe(0)
    expect(unmatched).toEqual(['Junction Improvement in Aleap Circle'])
    expect(out.rows[0]['Tender ID']).toBe('')
  })

  it('falls back to the embedding match when the exact name differs but scores above threshold', () => {
    const t = table([blankRow({ 'Name of the work': 'Junction Improvement Aleap Circle Ward 276' })])
    const { matchedCount } = updateWorksListFromEvaluations(t, [EV], {
      rowNameVectors: [[1, 0]],
      evalNameVectors: [[0.95, 0.05]]
    })
    expect(matchedCount).toBe(1)
  })

  it('does not use an embedding match below the threshold', () => {
    const t = table([blankRow({ 'Name of the work': 'Unrelated' })])
    const { matchedCount, unmatched } = updateWorksListFromEvaluations(t, [EV], {
      rowNameVectors: [[1, 0]],
      evalNameVectors: [[0, 1]]
    })
    expect(matchedCount).toBe(0)
    expect(unmatched).toEqual(['Junction Improvement in Aleap Circle'])
  })
})

describe('applyAgencyAddresses', () => {
  it('fills Address of the agency on every row of a matching agency, by agency name alone (not work name)', () => {
    const t = table([
      blankRow({ 'Name of the work': 'Work A', 'Name of the Agency': 'M V S CONSTRUCTIONS' }),
      blankRow({ 'Name of the work': 'Work B (different work, same agency)', 'Name of the Agency': 'M V S CONSTRUCTIONS' }),
      blankRow({ 'Name of the work': 'Work C', 'Name of the Agency': 'Other Builders' })
    ])
    const map = new Map([['m v s constructions', '13/B, Allwyn Colony, Hyderabad -500072']])
    const { table: out, filledCount } = applyAgencyAddresses(t, map)
    expect(filledCount).toBe(2)
    expect(out.rows[0]['Address of the agency']).toBe('13/B, Allwyn Colony, Hyderabad -500072')
    expect(out.rows[1]['Address of the agency']).toBe('13/B, Allwyn Colony, Hyderabad -500072')
    expect(out.rows[2]['Address of the agency']).toBe('') // agency not in the map
  })

  it('matches the agency name case/whitespace-insensitively and is idempotent', () => {
    const t = table([blankRow({ 'Name of the Agency': '  m v s   CONSTRUCTIONS ', 'Address of the agency': '13/B, Allwyn Colony' })])
    const map = new Map([['m v s constructions', '13/B, Allwyn Colony']])
    const { filledCount } = applyAgencyAddresses(t, map)
    expect(filledCount).toBe(0) // already the same value — no change
  })

  it('leaves rows with no agency name untouched', () => {
    const t = table([blankRow({ 'Name of the work': 'Work X', 'Name of the Agency': '' })])
    const map = new Map([['m v s constructions', 'somewhere']])
    const { filledCount } = applyAgencyAddresses(t, map)
    expect(filledCount).toBe(0)
  })
})

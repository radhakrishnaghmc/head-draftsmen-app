import { describe, expect, it } from 'vitest'
import { updateWorksListFromEvaluations } from '../core/worksTenderUpdate'
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
    const { table: out, matchedCount, unmatched, matchedRowIndices } = updateWorksListFromEvaluations(t, [EV])
    expect(matchedCount).toBe(0)
    expect(unmatched).toEqual(['Junction Improvement in Aleap Circle'])
    expect(matchedRowIndices).toEqual([])
    expect(out.rows[0]['Tender ID']).toBe('')
  })

  it('returns the matched row index so the caller can select the exact-matched row', () => {
    const t = table([
      blankRow({ 'Name of the work': 'Other work' }),
      blankRow({ 'Name of the work': 'Junction Improvement in Aleap Circle' })
    ])
    const { matchedRowIndices } = updateWorksListFromEvaluations(t, [EV])
    expect(matchedRowIndices).toEqual([1])
  })

  it('falls back to the embedding match when the exact name differs but scores above threshold', () => {
    const t = table([
      blankRow({ 'Name of the work': 'Unrelated first row' }),
      blankRow({ 'Name of the work': 'Junction Improvement Aleap Circle Ward 276' })
    ])
    const { matchedCount, matchedRowIndices } = updateWorksListFromEvaluations(t, [EV], {
      rowNameVectors: [[0, 1], [1, 0]],
      evalNameVectors: [[0.95, 0.05]]
    })
    expect(matchedCount).toBe(1)
    // The embedding match must report row 1 — this is the index the UI needs to
    // select, since there is no exact name to re-derive it from.
    expect(matchedRowIndices).toEqual([1])
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

  it('also fills EMD @ 1% / 1.5% and ASD from the ECV and Tender %', () => {
    const t = table([blankRow({ 'Name of the work': 'Junction Improvement in Aleap Circle' })])
    const r = updateWorksListFromEvaluations(t, [EV]).table.rows[0]
    expect(r['EMD 1%']).toBe('15935') // 1% of 1593493
    expect(r['EMD 1.5%']).toBe('23902') // 1.5% of 1593493
    expect(r['ASD']).toBe('0') // Tender % 11.11 is below the 25% ASD threshold
  })

  it('charges ASD at (Tender % - 25%) of ECV once Tender % exceeds 25%', () => {
    const t = table([blankRow({ 'Name of the work': 'Junction Improvement in Aleap Circle' })])
    const r = updateWorksListFromEvaluations(t, [{ ...EV, tenderPercentage: 30 }]).table.rows[0]
    expect(r['ASD']).toBe(String(Math.round(1593493 * 0.05))) // (30-25)% of ECV
  })

  it('folds the intimation address in, and uses it as a fallback for agency / contract', () => {
    const t = table([blankRow({ 'Name of the work': 'Junction Improvement in Aleap Circle' })])
    const notice = {
      agencyName: 'FALLBACK AGENCY',
      address: '12-3-45 Some Street, Hyderabad',
      contractRupees: 999999
    }
    // L-1 lacking agency & contract — the intimation supplies them; the address
    // (which only the intimation carries) is written regardless.
    const evNoAgency: TenderEvaluation = { ...EV, l1AgencyName: undefined, contractRupees: undefined }
    const r = updateWorksListFromEvaluations(t, [evNoAgency], undefined, notice).table.rows[0]
    expect(r['Address of the agency']).toBe('12-3-45 Some Street, Hyderabad')
    expect(r['Name of the Agency']).toBe('FALLBACK AGENCY')
    expect(r['Contract Amount']).toBe('999999')
  })

  it('prefers the L-1 sheet over the intimation for agency and contract when both are present', () => {
    const t = table([blankRow({ 'Name of the work': 'Junction Improvement in Aleap Circle' })])
    const notice = { agencyName: 'FALLBACK AGENCY', address: 'Addr', contractRupees: 999999 }
    const r = updateWorksListFromEvaluations(t, [EV], undefined, notice).table.rows[0]
    expect(r['Name of the Agency']).toBe('M V S CONSTRUCTIONS')
    expect(r['Contract Amount']).toBe('1416456')
    expect(r['Address of the agency']).toBe('Addr')
  })
})

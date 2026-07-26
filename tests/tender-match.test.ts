import { describe, expect, it } from 'vitest'
import { parseTenderRows, fillWorksListFromTenders } from '../core/tenderMatch'
import type { ExcelTable } from '../core/types'
import { WORKS_COLUMNS } from '../src/worksSchema'

function blankRow(overrides: Record<string, string> = {}): Record<string, string> {
  return { ...Object.fromEntries(WORKS_COLUMNS.map((h) => [h, ''])), ...overrides }
}

function table(rows: Record<string, string>[]): ExcelTable {
  return { id: 't1', name: 'Works database', path: '', headers: WORKS_COLUMNS, rows }
}

// A raw DataTables row shaped like the real portal response — only indices
// 0 (Organisation), 1 (Tender ID), 2 (NIT/Tender No), 4 (Work), 5 (ECV) are
// ever read.
function tenderRow(opts: { id?: string; notice?: string; work?: string; ecv?: string }): string[] {
  const r = new Array(10).fill('')
  r[0] = 'Some Circle'
  r[1] = opts.id ?? 'T-1'
  r[2] = opts.notice ?? 'NIT-1'
  r[4] = opts.work ?? 'Road work'
  r[5] = opts.ecv ?? '4500000'
  return r
}

describe('parseTenderRows', () => {
  it('extracts Tender ID / Notice No / Work name / ECV from the raw column positions', () => {
    const [t] = parseTenderRows([tenderRow({ id: 'T-99', notice: 'NIT-42', work: 'Road from A to B', ecv: '2500000' })])
    expect(t).toEqual({ tenderId: 'T-99', noticeNo: 'NIT-42', workName: 'Road from A to B', ecvRupees: 2500000 })
  })

  it('drops a row with no Tender ID or no Work name', () => {
    const rows = [tenderRow({ id: '' }), tenderRow({ work: '' }), tenderRow({})]
    expect(parseTenderRows(rows)).toHaveLength(1)
  })
})

describe('fillWorksListFromTenders', () => {
  it('fills ECV (rupees), Tender Notice No, and Tender ID for a row whose work name matches exactly', () => {
    const t = table([blankRow({ 'Name of the work': 'Road from A to B' })])
    const tenders = parseTenderRows([tenderRow({ id: 'T-1', notice: 'NIT-1', work: 'Road from A to B', ecv: '4500000' })])
    const { table: out, matchedCount } = fillWorksListFromTenders(t, tenders)

    expect(matchedCount).toBe(1)
    expect(out.rows[0]['ECV']).toBe('4500000')
    expect(out.rows[0]['Tender Notice No']).toBe('NIT-1')
    expect(out.rows[0]['Tender ID']).toBe('T-1')
  })

  it('is case/whitespace-insensitive for the exact match', () => {
    const t = table([blankRow({ 'Name of the work': '  road   from a to b  ' })])
    const tenders = parseTenderRows([tenderRow({ work: 'Road From A To B' })])
    const { matchedCount } = fillWorksListFromTenders(t, tenders)
    expect(matchedCount).toBe(1)
  })

  it('overwrites an existing ECV/Tender ID rather than only filling blanks', () => {
    const t = table([
      blankRow({ 'Name of the work': 'Road from A to B', ECV: '99', 'Tender ID': 'stale-id' })
    ])
    const tenders = parseTenderRows([tenderRow({ id: 'T-new', work: 'Road from A to B', ecv: '1000000' })])
    const { table: out } = fillWorksListFromTenders(t, tenders)
    expect(out.rows[0]['ECV']).toBe('1000000')
    expect(out.rows[0]['Tender ID']).toBe('T-new')
  })

  it('leaves a row untouched when no tender matches its work name, even below the embedding threshold', () => {
    const t = table([blankRow({ 'Name of the work': 'Completely unrelated work' })])
    const tenders = parseTenderRows([tenderRow({ work: 'Road from A to B' })])
    const embeddings = { rowNameVectors: [[1, 0]], tenderNameVectors: [[0, 1]] } // orthogonal -> score 0
    const { table: out, matchedCount } = fillWorksListFromTenders(t, tenders, embeddings)
    expect(matchedCount).toBe(0)
    expect(out.rows[0]['ECV']).toBe('')
  })

  it('falls back to the embedding match when the exact name differs but scores above threshold', () => {
    const t = table([blankRow({ 'Name of the work': 'Road from A to B (Phase 1)' })])
    const tenders = parseTenderRows([tenderRow({ id: 'T-2', work: 'Road from A to B Phase-1' })])
    const embeddings = { rowNameVectors: [[1, 0]], tenderNameVectors: [[0.9, 0.1]] } // high cosine similarity
    const { table: out, matchedCount } = fillWorksListFromTenders(t, tenders, embeddings)
    expect(matchedCount).toBe(1)
    expect(out.rows[0]['Tender ID']).toBe('T-2')
  })

  it('leaves a row with a blank "Name of the work" untouched', () => {
    const t = table([blankRow()])
    const tenders = parseTenderRows([tenderRow({})])
    const { matchedCount } = fillWorksListFromTenders(t, tenders)
    expect(matchedCount).toBe(0)
  })

  it('returns the table unchanged when there are no tenders to match against', () => {
    const t = table([blankRow({ 'Name of the work': 'Road from A to B' })])
    const { table: out, matchedCount } = fillWorksListFromTenders(t, [])
    expect(matchedCount).toBe(0)
    expect(out).toBe(t)
  })
})

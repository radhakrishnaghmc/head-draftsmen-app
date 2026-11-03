import { describe, it, expect } from 'vitest'
import { officeTabTitle, worksTableToGrid, gridSignature, chunkRows, columnLetters, quoteTab, worksGridsByOffice } from '@core/onlineWorksSheet'
import type { ExcelTable } from '@core/types'

const table = (rows: Record<string, string>[], extra: Partial<ExcelTable> = {}): ExcelTable => ({
  id: 't', name: 'Works', path: '', headers: ['A', 'B'], rows, ...extra
})

describe('online works sheet', () => {
  it('builds a legal tab title from the office key', () => {
    expect(officeTabTitle('CMC|Quthbullapur|Circle 1')).toBe('CMC - Quthbullapur - Circle 1')
    expect(officeTabTitle('X|a/b?c')).toBe('X - a b c')
    expect(officeTabTitle('').length).toBeGreaterThan(0)
    expect(officeTabTitle('x'.repeat(300)).length).toBe(100)
  })
  it('uses the source column order when known', () => {
    const g = worksTableToGrid(table([{ A: '1', B: '2' }], { sourceHeaders: ['B', 'A'] }))
    expect(g).toEqual([['B', 'A'], ['2', '1']])
    expect(worksTableToGrid(table([{ A: '1' }]))).toEqual([['A', 'B'], ['1', '']])
  })
  it('signature changes only when content changes', () => {
    const a = gridSignature([['a', 'b'], ['c']])
    expect(gridSignature([['a', 'b'], ['c']])).toBe(a)
    expect(gridSignature([['a', 'bc'], ['']])).not.toBe(a)
    expect(gridSignature([['a'], ['b', 'c']])).not.toBe(gridSignature([['a', 'b'], ['c']]))
  })
  it('chunks, letters and quotes', () => {
    expect(chunkRows([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]])
    expect([1, 26, 27, 52, 53].map(columnLetters)).toEqual(['A', 'Z', 'AA', 'AZ', 'BA'])
    expect(quoteTab("O'Neil")).toBe("'O''Neil'")
  })
  it('collects every office, current office winning', () => {
    const g = worksGridsByOffice({
      tablesByOffice: { 'A|B|C': [table([{ A: 'old' }])], 'D|E|F': [table([{ A: 'x' }])] },
      tables: [table([{ A: 'new' }])],
      currentOfficeKey: 'A|B|C'
    })
    expect(Object.keys(g).sort()).toEqual(['A|B|C', 'D|E|F'])
    expect(g['A|B|C'][1][0]).toBe('new')
  })
})

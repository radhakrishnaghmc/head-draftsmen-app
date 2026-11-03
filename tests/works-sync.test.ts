import { describe, it, expect } from 'vitest'
import { hasWorks, mergeNewWorks, scanBelongsToOffice, saveWorkspaceInputs } from '@core/worksSync'
import type { ExcelTable } from '@core/types'

const H = ['Wincode', 'Name of the work', 'ECV', 'Sanction By']
const t = (rows: Record<string, string>[]): ExcelTable => ({ id: 't', name: 'w', path: '', headers: H, rows })
const r = (Wincode: string, name: string, ECV = '', by = '') => ({ Wincode, 'Name of the work': name, ECV, 'Sanction By': by })

describe('works sync', () => {
  it('treats a blank starter row as no works', () => {
    expect(hasWorks(t([r('', '')]))).toBe(false)
    expect(hasWorks(t([r('D1', 'Laying CC road')]))).toBe(true)
  })
  it('re-import adds only new works and never updates existing ones', () => {
    const db = t([r('D1', 'Laying CC road at A', '500000')])
    const sheet = t([r('D1', 'Laying CC road at A (edited)', '1'), r('D2', 'BT road at B', '9'), r('', 'laying cc road at a', '2')])
    const { table, added, skipped } = mergeNewWorks(db, sheet)
    expect(added).toBe(1)
    expect(skipped).toBe(2)
    expect(table.rows.map((x) => x.Wincode)).toEqual(['D1', 'D2'])
    expect(table.rows[0].ECV).toBe('500000')
  })
  it('rejects another circle unless the work is already in the database', () => {
    const db = t([r('D1', 'Providing drain at Moosapet ward 10, Kukatpally Circle')])
    const base = { officeCircle: 'Gajularamaram', table: db }
    expect(scanBelongsToOffice({ ...base, sheetCircle: 'Gajularamaram', workName: 'x' })).toEqual({ ok: true })
    expect(scanBelongsToOffice({ ...base, sheetCircle: '', workName: 'x' })).toEqual({ ok: true })
    expect(scanBelongsToOffice({ ...base, sheetCircle: 'Nizampet', workName: 'Other work' })).toEqual({
      ok: false,
      error: 'These sheets do not belong to the Gajularamaram office.'
    })
    expect(scanBelongsToOffice({ ...base, sheetCircle: 'Kukatpally', workName: 'Providing drain at Moosapet ward 10, Kukatpally Circle' })).toEqual({ ok: true })
  })
  it('saves workspace inputs onto the row', () => {
    const db = { ...t([r('D1', 'a')]), headers: [...H, 'Newspapers Published'] }
    const out = saveWorkspaceInputs(db, 0, { newspapers: 'Eenadu', adminSanctionBy: 'Commissioner' })
    expect(out.rows[0]['Newspapers Published']).toBe('Eenadu')
    expect(out.rows[0]['Sanction By']).toBe('Commissioner')
    expect(saveWorkspaceInputs(out, 0, { newspapers: 'Eenadu' })).toBe(out)
  })
  it('refresh updates only Wincode, Eoffice and Amount of estimate on existing works', () => {
    const cols = ['Wincode', 'Eoffice', 'Name of the work', 'Amount of estimate', 'ECV']
    const mk = (rows: Record<string, string>[]): ExcelTable => ({ id: 't', name: 'w', path: '', headers: cols, rows })
    const db = mk([{ Wincode: 'D1', Eoffice: '100', 'Name of the work': 'Road A', 'Amount of estimate': '10', ECV: '900' }])
    const sheet = mk([
      { Wincode: 'D1', Eoffice: '200', 'Name of the work': 'Road A', 'Amount of estimate': '12', ECV: '1' },
      { Wincode: 'D2', Eoffice: '300', 'Name of the work': 'Road B', 'Amount of estimate': '5', ECV: '' }
    ])
    const out = mergeNewWorks(db, sheet)
    expect(out.updated).toBe(1)
    expect(out.added).toBe(1)
    expect(out.table.rows[0]).toMatchObject({ Eoffice: '200', 'Amount of estimate': '12', ECV: '900' })
    expect(out.table.rows[1].Eoffice).toBe('300')
    expect(mergeNewWorks(out.table, sheet).updated).toBe(0)
  })
})

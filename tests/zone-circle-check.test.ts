import { describe, expect, it } from 'vitest'
import { enforceZoneCircle, fillCircleNumber } from '../src/zoneCircleCheck'
import type { ExcelTable } from '../core/types'

function table(rows: Record<string, string>[], headers = ['Zone', 'Circle', 'Name of the work']): ExcelTable {
  return {
    id: 't1',
    name: 'Works database',
    path: '',
    headers,
    rows
  }
}

describe('enforceZoneCircle', () => {
  it('passes rows whose explicit Zone/Circle already match the login identity', () => {
    const t = table([{ Zone: 'Cyberabad', Circle: 'Gajularamaram Circle-57', 'Name of the work': 'Road work' }])
    const result = enforceZoneCircle(t, 'Cyberabad', 'Gajularamaram Circle-57')
    expect(result.mismatches).toHaveLength(0)
    expect(result.filledCount).toBe(0)
  })

  it('matches case/whitespace-insensitively, not just exact string equality', () => {
    const t = table([{ Zone: '  cyberabad ', Circle: 'GAJULARAMARAM CIRCLE-57', 'Name of the work': 'Road work' }])
    const result = enforceZoneCircle(t, 'Cyberabad', 'Gajularamaram Circle-57')
    expect(result.mismatches).toHaveLength(0)
  })

  it('flags a row whose explicit Zone/Circle conflicts with the login identity', () => {
    const t = table([{ Zone: 'Warangal', Circle: 'Some Other Circle-1', 'Name of the work': 'Road work' }])
    const result = enforceZoneCircle(t, 'Cyberabad', 'Gajularamaram Circle-57')
    expect(result.mismatches).toHaveLength(1)
    expect(result.mismatches[0]).toMatchObject({ foundZone: 'Warangal', foundCircle: 'Some Other Circle-1' })
  })

  it('fills a blank Zone/Circle when the login identity is mentioned in the work name', () => {
    const t = table([
      { Zone: '', Circle: '', 'Name of the work': 'Road work in Gajularamaram Circle-57, Cyberabad zone' }
    ])
    const result = enforceZoneCircle(t, 'Cyberabad', 'Gajularamaram Circle-57')
    expect(result.mismatches).toHaveLength(0)
    expect(result.filledCount).toBe(2)
    expect(result.table.rows[0].Zone).toBe('Cyberabad')
    expect(result.table.rows[0].Circle).toBe('Gajularamaram Circle-57')
  })

  it('leaves a row alone when Zone/Circle are blank and nothing in the name can confirm or deny it', () => {
    const t = table([{ Zone: '', Circle: '', 'Name of the work': 'Generic road work' }])
    const result = enforceZoneCircle(t, 'Cyberabad', 'Gajularamaram Circle-57')
    expect(result.mismatches).toHaveLength(0)
    expect(result.filledCount).toBe(0)
    expect(result.table.rows[0].Zone).toBe('')
  })

  it('only reports the rows that actually conflict, in a multi-row table', () => {
    const t = table([
      { Zone: 'Cyberabad', Circle: 'Gajularamaram Circle-57', 'Name of the work': 'Work A' },
      { Zone: 'Warangal', Circle: 'Gajularamaram Circle-57', 'Name of the work': 'Work B' },
      { Zone: '', Circle: '', 'Name of the work': 'Work C in Cyberabad zone' }
    ])
    const result = enforceZoneCircle(t, 'Cyberabad', 'Gajularamaram Circle-57')
    expect(result.mismatches).toHaveLength(1)
    expect(result.mismatches[0].rowIndex).toBe(1)
    expect(result.table.rows[2].Zone).toBe('Cyberabad')
  })
})

describe('fillCircleNumber', () => {
  it('fills a blank CNO cell with the login circle number', () => {
    const t = table([{ CNO: '', Zone: 'Cyberabad', Circle: 'Gajularamaram Circle-57' }], ['Zone', 'Circle', 'CNO'])
    const result = fillCircleNumber(t, '57')
    expect(result.rows[0].CNO).toBe('57')
  })

  it('leaves an already-filled CNO cell untouched', () => {
    const t = table([{ CNO: '12', Zone: 'Cyberabad', Circle: 'Gajularamaram Circle-57' }], ['Zone', 'Circle', 'CNO'])
    const result = fillCircleNumber(t, '57')
    expect(result.rows[0].CNO).toBe('12')
  })

  it('does nothing when no circle number is known (e.g. a Zone-level login)', () => {
    const t = table([{ CNO: '', Zone: 'Cyberabad', Circle: '' }], ['Zone', 'Circle', 'CNO'])
    const result = fillCircleNumber(t, undefined)
    expect(result.rows[0].CNO).toBe('')
  })

  it('does nothing when the table has no CNO column', () => {
    const t = table([{ Zone: 'Cyberabad', Circle: 'Gajularamaram Circle-57' }])
    const result = fillCircleNumber(t, '57')
    expect(result).toEqual(t)
  })
})

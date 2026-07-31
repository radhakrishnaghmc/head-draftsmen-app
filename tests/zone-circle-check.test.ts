import { describe, expect, it } from 'vitest'
import { autofillWorksRow, enforceZoneCircle, fillCircleNumber, splitCircleColumn } from '../src/zoneCircleCheck'
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

  // The whole list is the logged-in Head Draughtsman's own circle, so a blank row
  // whose name mentions no other circle is stamped with the login identity —
  // work names usually mention neither their circle nor their zone.
  it('fills blank Zone/Circle with the login identity when the name names no other circle', () => {
    const t = table([{ Zone: '', Circle: '', 'Name of the work': 'Improvements to Peddamma temple road' }])
    const result = enforceZoneCircle(t, 'Quthbullapur', 'Gajularamaram')
    expect(result.mismatches).toHaveLength(0)
    expect(result.filledCount).toBe(2)
    expect(result.table.rows[0].Zone).toBe('Quthbullapur')
    expect(result.table.rows[0].Circle).toBe('Gajularamaram')
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

  // A work name names its circle but not its zone — the CMC zone↔circle
  // directory infers the zone so both blank columns get filled.
  it('infers the zone from the circle named in the work name and fills both columns', () => {
    const t = table([{ Zone: '', Circle: '', 'Name of the work': 'Improvements to CC road at Miyapur' }])
    const result = enforceZoneCircle(t, 'Serilingampally', 'Miyapur')
    expect(result.mismatches).toHaveLength(0)
    expect(result.filledCount).toBe(2)
    expect(result.table.rows[0].Circle).toBe('Miyapur')
    expect(result.table.rows[0].Zone).toBe('Serilingampally')
  })

  it('flags a work whose name names a circle belonging to a different zone', () => {
    const t = table([{ Zone: '', Circle: '', 'Name of the work': 'Road work at Kompally' }])
    const result = enforceZoneCircle(t, 'Serilingampally', 'Miyapur')
    expect(result.mismatches).toHaveLength(1)
    expect(result.mismatches[0]).toMatchObject({ foundCircle: 'Kompally', foundZone: 'Quthbullapur' })
  })

  // A road running between two circles in different zones is ambiguous — the
  // directory must not guess a zone from it, but the login circle appearing in
  // the name still fills the row's own identity.
  it('does not guess when the name spans two circles in different zones', () => {
    const t = table([{ Zone: '', Circle: '', 'Name of the work': 'CC road from Miyapur to Nizampet' }])
    const result = enforceZoneCircle(t, 'Serilingampally', 'Miyapur')
    expect(result.mismatches).toHaveLength(0)
    expect(result.table.rows[0].Circle).toBe('Miyapur')
    expect(result.table.rows[0].Zone).toBe('Serilingampally')
  })
})

describe('fillCircleNumber', () => {
  it('fills a blank Circle number cell with the login circle number', () => {
    const t = table([{ 'Circle number': '', Zone: 'Cyberabad', Circle: 'Gajularamaram' }], [
      'Zone',
      'Circle',
      'Circle number'
    ])
    const result = fillCircleNumber(t, '57')
    expect(result.rows[0]['Circle number']).toBe('57')
  })

  it('still recognizes the legacy "CNO" column name', () => {
    const t = table([{ CNO: '', Zone: 'Cyberabad', Circle: 'Gajularamaram' }], ['Zone', 'Circle', 'CNO'])
    const result = fillCircleNumber(t, '57')
    expect(result.rows[0].CNO).toBe('57')
  })

  it('leaves an already-filled Circle number cell untouched', () => {
    const t = table([{ 'Circle number': '12', Zone: 'Cyberabad', Circle: 'Gajularamaram' }], [
      'Zone',
      'Circle',
      'Circle number'
    ])
    const result = fillCircleNumber(t, '57')
    expect(result.rows[0]['Circle number']).toBe('12')
  })

  it('does nothing when no circle number is known (e.g. a Zone-level login)', () => {
    const t = table([{ 'Circle number': '', Zone: 'Cyberabad', Circle: '' }], ['Zone', 'Circle', 'Circle number'])
    const result = fillCircleNumber(t, undefined)
    expect(result.rows[0]['Circle number']).toBe('')
  })

  it('does nothing when the table has no Circle number column', () => {
    const t = table([{ Zone: 'Cyberabad', Circle: 'Gajularamaram' }])
    const result = fillCircleNumber(t, '57')
    expect(result).toEqual(t)
  })
})

describe('splitCircleColumn', () => {
  const cols = ['Zone', 'Circle', 'Circle number', 'Name of the work']

  it('moves a number embedded in the Circle cell into the Circle number column', () => {
    const t = table([{ Zone: '', Circle: '57-Gajularamaram', 'Circle number': '', 'Name of the work': 'x' }], cols)
    const result = splitCircleColumn(t)
    expect(result.rows[0].Circle).toBe('Gajularamaram')
    expect(result.rows[0]['Circle number']).toBe('57')
  })

  it('handles the number written after the name ("Gajularamaram 57")', () => {
    const t = table([{ Zone: '', Circle: 'Gajularamaram 57', 'Circle number': '', 'Name of the work': 'x' }], cols)
    const result = splitCircleColumn(t)
    expect(result.rows[0].Circle).toBe('Gajularamaram')
    expect(result.rows[0]['Circle number']).toBe('57')
  })

  it('leaves a bare circle name and its own number untouched', () => {
    const t = table([{ Zone: '', Circle: 'Gajularamaram', 'Circle number': '57', 'Name of the work': 'x' }], cols)
    const result = splitCircleColumn(t)
    expect(result.rows[0].Circle).toBe('Gajularamaram')
    expect(result.rows[0]['Circle number']).toBe('57')
  })

  it('does not clobber an already-filled Circle number', () => {
    const t = table([{ Zone: '', Circle: '57-Gajularamaram', 'Circle number': '99', 'Name of the work': 'x' }], cols)
    const result = splitCircleColumn(t)
    expect(result.rows[0].Circle).toBe('Gajularamaram')
    expect(result.rows[0]['Circle number']).toBe('99')
  })
})

describe('autofillWorksRow (live in-grid auto-fill)', () => {
  const blank = (name: string) => ({ Zone: '', Circle: '', 'Circle number': '', 'Name of the work': name })

  it('fills Zone/Circle/Circle number from a work name tagged with its circle and zone', () => {
    const filled = autofillWorksRow(
      blank('laying of miyapur to bachupally road in bachupally ward of nizampet circle, quthbullapur zone cmc'),
      { zone: 'Quthbullapur', circle: 'Nizampet', circleNumber: '58' }
    )
    expect(filled).toMatchObject({ Zone: 'Quthbullapur', Circle: 'Nizampet', 'Circle number': '58' })
  })

  it('fills from a work name that only tags the circle (zone via directory)', () => {
    const filled = autofillWorksRow(blank('laying of cc road in kphb to jntu in kukatpally circle, cmc'), {
      zone: 'Kukatpally',
      circle: 'Kukatpally',
      circleNumber: '52'
    })
    expect(filled).toMatchObject({ Zone: 'Kukatpally', Circle: 'Kukatpally', 'Circle number': '52' })
  })

  it('splits a combined Circle cell the user typed and fills the number', () => {
    const filled = autofillWorksRow(
      { Zone: '', Circle: '57-Gajularamaram', 'Circle number': '', 'Name of the work': '' },
      { zone: 'Quthbullapur', circle: 'Gajularamaram', circleNumber: '57' }
    )
    expect(filled).toMatchObject({ Circle: 'Gajularamaram', 'Circle number': '57', Zone: 'Quthbullapur' })
  })

  it('never overwrites values the user already entered', () => {
    const filled = autofillWorksRow(
      { Zone: 'Manual zone', Circle: 'Manual circle', 'Circle number': '99', 'Name of the work': 'road at Miyapur' },
      { zone: 'Serilingampally', circle: 'Miyapur', circleNumber: '48' }
    )
    expect(filled).toMatchObject({ Zone: 'Manual zone', Circle: 'Manual circle', 'Circle number': '99' })
  })
})

import { describe, it, expect } from 'vitest'
import { checkLoaAgainstL1, pairForIntimations } from '../core/bulkIntimations'

const l1 = { nameOfWork: 'Road A', noticeNo: '12/EE/Circle-5/2026-27', l1AgencyName: 'M/s ABC Constructions', contractRupees: 500000.5 } as any
const loa = { nitNo: '12/EE/Circle-5/2026-27 Dated 01.01.2026', agencyName: 'ABC CONSTRUCTIONS', contractRupees: 500000.5 } as any

describe('checkLoaAgainstL1', () => {
  it('accepts a matching LOA', () => expect(checkLoaAgainstL1(loa, l1).ok).toBe(true))
  it('rejects a different contract amount', () => {
    const r = checkLoaAgainstL1({ ...loa, contractRupees: 600000 }, l1)
    expect(r.ok).toBe(false)
    expect(r.problems[0]).toMatch(/Contract amount differs/)
  })
  it('rejects a different agency and a different NIT', () => {
    const r = checkLoaAgainstL1({ ...loa, agencyName: 'XYZ Builders', nitNo: '99/EE/Circle-9/2026-27' }, l1)
    expect(r.problems.length).toBe(2)
  })
  it('rejects a missing contract amount', () => expect(checkLoaAgainstL1({ ...loa, contractRupees: undefined }, l1).ok).toBe(false))
})

describe('pairForIntimations', () => {
  it('pairs a good LOA and reports a bad one with its reason', () => {
    const bad = { ...l1, noticeNo: '77/EE/Circle-5/2026-27', l1AgencyName: 'Other Agency', contractRupees: 10 }
    const badLoa = { ...loa, nitNo: '77/EE/Circle-5/2026-27', agencyName: 'Other Agency', contractRupees: 500 }
    const r = pairForIntimations(
      [{ file: 'a.pdf', doc: l1 }, { file: 'b.pdf', doc: bad }],
      [{ file: 'a.html', doc: loa }, { file: 'b.html', doc: badLoa }]
    )
    expect(r.pairs.length).toBe(1)
    expect(r.rejected.length).toBe(1)
    expect(r.rejected[0].reason).toMatch(/Contract amount differs/)
  })
  it('says so when no LOA is related', () => {
    const r = pairForIntimations([{ file: 'a.pdf', doc: l1 }], [])
    expect(r.rejected[0].reason).toMatch(/No LOA/)
  })
})

import { describe, it, expect } from 'vitest'
import { mergeQcParties } from '../core/mergeQcParties'

const party = (name: string) => ({ name, address: '', phone: '' })
const empty = party('')

describe('mergeQcParties', () => {
  it('keeps local agencies the cloud does not have yet', () => {
    const local = { A: { third: party('College'), fourth: party('Lab') } }
    expect(mergeQcParties({}, local)).toEqual(local)
    expect(mergeQcParties(undefined, local)).toEqual(local)
  })
  it('keeps local agencies when the cloud entry is blank', () => {
    const out = mergeQcParties({ A: { third: empty, fourth: empty } }, { A: { third: party('College'), fourth: party('Lab') } })
    expect(out!.A.third.name).toBe('College')
    expect(out!.A.fourth.name).toBe('Lab')
  })
  it('prefers the cloud where it has details, and keeps other offices from both sides', () => {
    const out = mergeQcParties(
      { A: { third: party('Cloud'), fourth: empty }, B: { third: party('B3'), fourth: empty } },
      { A: { third: party('Local'), fourth: party('Lab') }, C: { third: party('C3'), fourth: empty } }
    )
    expect(out!.A.third.name).toBe('Cloud')
    expect(out!.A.fourth.name).toBe('Lab')
    expect(out!.B.third.name).toBe('B3')
    expect(out!.C.third.name).toBe('C3')
  })
  it('stays undefined when neither side has any', () => {
    expect(mergeQcParties(undefined, undefined)).toBeUndefined()
  })
})

describe('mergeQcParties department QC', () => {
  it('keeps a local department QC the cloud lacks, and prefers the cloud one when present', () => {
    const local = { A: { department: party('Dept'), third: empty, fourth: empty } }
    expect(mergeQcParties({ A: { third: empty, fourth: empty } }, local)!.A.department!.name).toBe('Dept')
    expect(mergeQcParties({ A: { department: party('Cloud'), third: empty, fourth: empty } }, local)!.A.department!.name).toBe('Cloud')
  })
})

import { describe, expect, it } from 'vitest'
import { checkSameWork } from '../core/sameWorkCheck'
import type { IntimationNotice } from '../core/intimationNotice'
import type { TenderEvaluation } from '../core/tenderEvaluationPdf'

const notice = (o: Partial<IntimationNotice>): IntimationNotice => o
const pdf = (o: Partial<TenderEvaluation>): TenderEvaluation => o

describe('checkSameWork', () => {
  it('matches when both NIT Nos are equal (ignoring spacing around / and -)', () => {
    const r = checkSameWork(
      notice({ nitNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27' }),
      pdf({ noticeNo: '12/DB/EE/Nizampet Circle- 58/CMC/2026-27', nameOfWork: 'X' })
    )
    expect(r.status).toBe('match')
    expect(r.by).toBe('nit')
  })

  it('flags a mismatch when the NIT Nos differ', () => {
    const r = checkSameWork(
      notice({ nitNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27' }),
      pdf({ noticeNo: '99/DB/EE/Other Circle-1/CMC/2026-27', nameOfWork: 'Y' })
    )
    expect(r.status).toBe('mismatch')
    expect(r.by).toBe('nit')
  })

  it('matches with the agency taking priority when both the agency and NIT No agree (M/s prefix ignored)', () => {
    const r = checkSameWork(
      notice({ nitNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27', agencyName: 'M V S CONSTRUCTIONS' }),
      pdf({ noticeNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27', l1AgencyName: 'M/s M V S Constructions', nameOfWork: 'X' })
    )
    expect(r.status).toBe('match')
    expect(r.by).toBe('agency')
  })

  it('flags a mismatch on the agency even when the NIT Nos agree (agency has priority)', () => {
    const r = checkSameWork(
      notice({ nitNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27', agencyName: 'Kummary Renuka Devi Civil Contractor' }),
      pdf({ noticeNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27', l1AgencyName: 'SRI SHARVA CONSTRUCTIONS', nameOfWork: 'X' })
    )
    expect(r.status).toBe('mismatch')
    expect(r.by).toBe('agency')
  })

  it('flags a mismatch when the NIT Nos match but the L-1 agency is a different firm', () => {
    const r = checkSameWork(
      notice({ nitNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27', agencyName: 'M V S CONSTRUCTIONS' }),
      pdf({ noticeNo: '12/DB/EE/Nizampet Circle-58/CMC/2026-27', l1AgencyName: 'ABC INFRA PROJECTS', nameOfWork: 'X' })
    )
    expect(r.status).toBe('mismatch')
    expect(r.by).toBe('agency')
  })

  it('falls back to agency name when a NIT No is missing on one side', () => {
    const same = checkSameWork(
      notice({ agencyName: 'M V S CONSTRUCTIONS' }),
      pdf({ l1AgencyName: 'M V S Constructions', nameOfWork: 'Z' })
    )
    expect(same.status).toBe('match')
    expect(same.by).toBe('agency')

    const diff = checkSameWork(
      notice({ agencyName: 'M V S CONSTRUCTIONS' }),
      pdf({ l1AgencyName: 'ABC INFRA', nameOfWork: 'Z' })
    )
    expect(diff.status).toBe('mismatch')
    expect(diff.by).toBe('agency')
  })

  it('is unknown when neither identifier is available on both sides', () => {
    expect(checkSameWork(notice({}), pdf({ nameOfWork: 'W' })).status).toBe('unknown')
    expect(checkSameWork(notice({ agencyName: 'A' }), pdf({ noticeNo: '1' })).status).toBe('unknown')
  })

  it("matches when the Tender Document threads 'Enquiry/IFB/Tender' through the NIT No the Intimation omits", () => {
    // Same work: the Online Intimation prints ".../C-55/..." while the full
    // Tender Document prints ".../C-Enquiry/IFB/Tender 55/...".
    const r = checkSameWork(
      notice({ nitNo: '05/DB/EE-III/JDM/C-55/QBZ/CMC/2026-27' }),
      pdf({ noticeNo: '05/DB/EE-III/JDM/C-Enquiry/IFB/Tender 55/QBZ/CMC/2026-27', nameOfWork: 'X' })
    )
    expect(r.status).toBe('match')
    expect(r.by).toBe('nit')

    // The serial still has to agree — a different number is not the same work.
    const diff = checkSameWork(
      notice({ nitNo: '05/DB/EE-III/JDM/C-55/QBZ/CMC/2026-27' }),
      pdf({ noticeNo: '05/DB/EE-III/JDM/C-Enquiry/IFB/Tender 56/QBZ/CMC/2026-27', nameOfWork: 'X' })
    )
    expect(diff.status).toBe('mismatch')
    expect(diff.by).toBe('nit')
  })

  it("matches when the Intimation's NIT No carries a trailing 'Dated …' the L-1's omits", () => {
    const r = checkSameWork(
      notice({ nitNo: '13/DB/EE/Nizampet Circle-58/CMC/2026-27 Dated 24.07.2026' }),
      pdf({ noticeNo: '13/DB/EE/Nizampet Circle-58/CMC/2026-27', nameOfWork: 'W' })
    )
    expect(r.status).toBe('match')
    expect(r.by).toBe('nit')

    // A genuinely different NIT No is still a mismatch.
    const diff = checkSameWork(
      notice({ nitNo: '12/DB/EE/Kompally Circle-56/CMC/2026-27 Dated 24.07.2026' }),
      pdf({ noticeNo: '13/DB/EE/Nizampet Circle-58/CMC/2026-27', nameOfWork: 'W' })
    )
    expect(diff.status).toBe('mismatch')
    expect(diff.by).toBe('nit')
  })
})

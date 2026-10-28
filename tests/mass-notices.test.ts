import { describe, expect, it } from 'vitest'
import { duplicateNoticeIds, pairL1WithLoa, hasAgreementDate, findMonitoringWork, decideFromMonitoringFormat } from '../core/massNotices'
import type { MonitoringFormatWorkRow } from '../core/monitoringFormat'

const l1 = (nit: string, agency: string, work: string, contract?: number) => ({
  file: `l1-${nit}.pdf`,
  doc: { noticeNo: nit, l1AgencyName: agency, nameOfWork: work, contractRupees: contract }
})
const loa = (nit: string, agency: string, contract?: number) => ({
  file: `loa-${nit}.html`,
  doc: { nitNo: nit, agencyName: agency, contractRupees: contract }
})

describe('pairL1WithLoa', () => {
  it('pairs an L-1 with the LOA of the same agency and NIT', () => {
    const r = pairL1WithLoa(
      [l1('NIT/55/2026-27', 'M/s AHR Constructions', 'CC road at X')],
      [loa('NIT/55/2026-27', 'AHR CONSTRUCTIONS')]
    )
    expect(r.pairs).toHaveLength(1)
    expect(r.noLoa).toHaveLength(0)
  })

  it('gives no notice when the L-1 has no LOA', () => {
    const r = pairL1WithLoa([l1('NIT/55/2026-27', 'AHR Constructions', 'CC road at X')], [])
    expect(r.pairs).toHaveLength(0)
    expect(r.noLoa).toHaveLength(1)
  })

  it('does not pair an LOA of a different work (different NIT, or different agency)', () => {
    const l = [l1('NIT/55/2026-27', 'AHR Constructions', 'CC road at X')]
    expect(pairL1WithLoa(l, [loa('NIT/56/2026-27', 'AHR Constructions')]).pairs).toHaveLength(0)
    expect(pairL1WithLoa(l, [loa('NIT/55/2026-27', 'Sri Sai Builders')]).pairs).toHaveLength(0)
  })

  it('uses each LOA once, preferring the one whose contract value agrees', () => {
    const r = pairL1WithLoa(
      [l1('NIT/9/2026-27', 'ABC Infra', 'Item 1', 500000), l1('NIT/9/2026-27', 'ABC Infra', 'Item 2', 900000)],
      [loa('NIT/9/2026-27', 'ABC Infra', 900000), loa('NIT/9/2026-27', 'ABC Infra', 500000)]
    )
    expect(r.pairs.map((p) => p.loa.contractRupees)).toEqual([500000, 900000])
  })
})

describe('pairL1WithLoa — Supplier Details fallback', () => {
  const l1WithId = (id: string, agency: string) => ({
    file: 'l1.pdf',
    doc: { tenderId: id, noticeNo: 'NIT/55/2026-27', l1AgencyName: agency, nameOfWork: 'CC road at X' }
  })
  const sup = (id: string, agency: string) => ({ file: 'supplier.pdf', doc: { tenderId: id, agencyName: agency, phone: '9666054562' } })

  it('accepts the L-1 when there is no LOA but the Supplier Details Tender ID and agency match', () => {
    const r = pairL1WithLoa([l1WithId('728662', 'M/s Lakshmi Constructions')], [], [sup('728662', 'LAKSHMI CONSTRUCTIONS')])
    expect(r.pairs).toHaveLength(1)
    expect(r.pairs[0].via).toBe('supplier')
    expect(r.pairs[0].loa.nitNo).toBe('NIT/55/2026-27')
    expect(r.pairs[0].loa.phone).toBe('9666054562')
  })

  it('rejects a different Tender ID or a different agency', () => {
    expect(pairL1WithLoa([l1WithId('728662', 'Lakshmi Constructions')], [], [sup('728663', 'Lakshmi Constructions')]).pairs).toHaveLength(0)
    expect(pairL1WithLoa([l1WithId('728662', 'Lakshmi Constructions')], [], [sup('728662', 'Sri Sai Builders')]).noLoa).toHaveLength(1)
  })

  it('prefers the LOA when both exist', () => {
    const r = pairL1WithLoa(
      [l1WithId('728662', 'Lakshmi Constructions')],
      [loa('NIT/55/2026-27', 'Lakshmi Constructions')],
      [sup('728662', 'Lakshmi Constructions')]
    )
    expect(r.pairs[0].via).toBe('loa')
  })
})

describe('hasAgreementDate', () => {
  it.each([
    '12.03.2026',
    '12/03/2026',
    '12-3-26',
    '2026-03-12',
    '12-Mar-2026',
    '12th March 2026',
    'March 12, 2026',
    '46093'
  ])('treats %s as a date', (s) => {
    expect(hasAgreementDate(s)).toBe(true)
  })
  it.each(['', '   ', 'Not yet', 'Under process', 'Pending', 'LOA issued', 'Yet to conclude, LOA dt 12.03.2026', 'NA'])(
    'treats %j as no date',
    (s) => {
      expect(hasAgreementDate(s)).toBe(false)
    }
  )
})

describe('Monitoring Format decision', () => {
  const row = (workName: string, agreementDate: string) => ({ workName, agreementDate }) as MonitoringFormatWorkRow
  const rows = [
    row('Construction of CC road at Ward 5', '12.03.2026'),
    row('Laying of SWD at Ward 7', 'Not yet'),
    row('Providing CC drain at Ward 9', '')
  ]

  it('skips when the agreement date is present', () => {
    const m = findMonitoringWork(rows, 'Construction of CC road at Ward 5')
    expect(decideFromMonitoringFormat(m, rows).kind).toBe('skip')
  })
  it('generates when the cell holds words or is blank', () => {
    expect(decideFromMonitoringFormat(findMonitoringWork(rows, 'Laying of SWD at Ward 7'), rows).kind).toBe('generate')
    expect(decideFromMonitoringFormat(findMonitoringWork(rows, 'Providing CC drain at Ward 9'), rows).kind).toBe(
      'generate'
    )
  })
  it('generates when the work is not in the sheet, or no sheet is imported', () => {
    expect(decideFromMonitoringFormat(findMonitoringWork(rows, 'Something else entirely'), rows).kind).toBe('generate')
    expect(decideFromMonitoringFormat(undefined, undefined).kind).toBe('generate')
  })
})

describe('duplicateNoticeIds', () => {
  const n = (id: string, o: object = {}) => ({ id, done: false, ...o })
  it('flags repeats of the same Tender ID, keeping the first', () => {
    expect(duplicateNoticeIds([n('a', { tenderId: '741316' }), n('b', { tenderId: '741316' }), n('c', { tenderId: '741317' })])).toEqual(['b'])
  })
  it('keeps a resolved or edited notice over a plain one', () => {
    expect(duplicateNoticeIds([n('a', { tenderId: '1' }), n('b', { tenderId: '1', done: true })])).toEqual(['a'])
    expect(duplicateNoticeIds([n('a', { tenderId: '1' }), n('b', { tenderId: '1', edits: { x: 'y' } })])).toEqual(['a'])
  })
  it('falls back to NIT + agency + work when there is no Tender ID, and ignores deleted ones', () => {
    const w = { nitNo: '27/2/DB', agencyName: 'AHR Constructions', workName: 'CC road' }
    expect(duplicateNoticeIds([n('a', w), n('b', w), n('c', { ...w, nitNo: '27/3/DB' }), n('d', { ...w, deletedAt: 'x' })])).toEqual(['b'])
  })
})

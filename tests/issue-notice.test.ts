import { describe, it, expect } from 'vitest'
import { issueNoticePlaceholders, EMPTY_ISSUE_NOTICE_MANUAL_FIELDS } from '../core/issueNotice'
import type { IntimationNotice } from '../core/intimationNotice'
import type { TenderEvaluation } from '../core/tenderEvaluationPdf'

describe('issueNoticePlaceholders', () => {
  const notice: IntimationNotice = {
    agencyName: 'M/s. Jaishvini Constructions',
    address: 'Villa No.22, EIPL riveredge, Manchirevula, Rangareddy, Telangana-500089',
    nitNo: '11/SE/QBZ/CMC/2026-27',
    nitDate: '25.06.2026',
    ecvRupees: 22723948
  }
  const pdf: TenderEvaluation = {
    nameOfWork: 'Widening of BT Road from Bowrampet Main Road Connecting to ORR Service Road in Bowrampet ward-295 in Dundigal Circle-59, Quthbullapur Zone, CMC (Item No.1)',
    tenderPercentage: 0.1,
    contractRupees: 22710224.05
  }
  const manual = { ...EMPTY_ISSUE_NOTICE_MANUAL_FIELDS, loaNo: '27', loaDate: '01.07.2026', noticeDate: '01.08.2026' }

  it('fills the amounts using the SE-letter Indian-grouped formatting, ECV from the Online Intimation', () => {
    const p = issueNoticePlaceholders(notice, pdf, manual, { zone: 'Quthbullapur' })
    expect(p['Tender Contract Value']).toBe('2,27,10,224.05')
    expect(p['Estimate Contract Value']).toBe('2,27,23,948.00')
    expect(p['Tender Contract Value In Words']).toContain('and')
  })

  it('derives Zone Abbr and strips the Item No tag from the work name into its own field', () => {
    const p = issueNoticePlaceholders(notice, pdf, manual, { zone: 'Quthbullapur' })
    expect(p['Zone Abbr']).toBe('QBZ')
    expect(p['Item No']).toBe('1')
    expect(p['Name of the work']).not.toMatch(/Item No/i)
  })

  it('NIT No/Date come from the Online Intimation first, falling back to the L1 sheet', () => {
    const p = issueNoticePlaceholders(notice, pdf, manual, { zone: 'Quthbullapur' })
    expect(p['Nit No']).toBe('11/SE/QBZ/CMC/2026-27')
    expect(p['Nit Date']).toBe('25.06.2026')

    const noNotice: IntimationNotice = {}
    const pdfOnly: TenderEvaluation = { ...pdf, noticeNo: '10/SE/QBZ/CMC/2026-27', noticeDate: '18.06.2026' }
    const p2 = issueNoticePlaceholders(noNotice, pdfOnly, manual, { zone: 'Quthbullapur' })
    expect(p2['Nit No']).toBe('10/SE/QBZ/CMC/2026-27')
    expect(p2['Nit Date']).toBe('18.06.2026')
  })

  it('LOA No/Date and Notice Date come only from the manually-typed fields — never derived', () => {
    const p = issueNoticePlaceholders(notice, pdf, manual, { zone: 'Quthbullapur' })
    expect(p['LOA No']).toBe('27')
    expect(p['LOA Date']).toBe('01.07.2026')
    expect(p['Notice Date']).toBe('01.08.2026')
  })

  it('Financial year derives from the Notice Date, falling back to the current FY when blank', () => {
    const p = issueNoticePlaceholders(notice, pdf, manual, { zone: 'Quthbullapur' })
    expect(p['Financial year']).toBe('2026-27')

    const noDate = issueNoticePlaceholders(notice, pdf, EMPTY_ISSUE_NOTICE_MANUAL_FIELDS, { zone: 'Quthbullapur' })
    expect(noDate['Financial year']).toMatch(/^\d{4}-\d{2}$/)
  })

  it('leaves amount placeholders blank rather than "0.00" when neither upload carries the figures', () => {
    const p = issueNoticePlaceholders({}, {}, EMPTY_ISSUE_NOTICE_MANUAL_FIELDS, { zone: 'Quthbullapur' })
    expect(p['Tender Contract Value']).toBe('')
    expect(p['Estimate Contract Value']).toBe('')
    expect(p['Tender Contract Value In Words']).toBe('')
    expect(p['Tender Percentage']).toBe('')
  })

  it('fills Circle/CNO for an EE (circle) office and leaves Zone/Zone Abbr blank, and vice versa for an SE (zone) office', () => {
    const ee = issueNoticePlaceholders(notice, pdf, manual, { circle: 'Gajularamaram', cno: '57' })
    expect(ee['Circle']).toBe('Gajularamaram')
    expect(ee['CNO']).toBe('57')
    expect(ee['Zone']).toBe('')
    expect(ee['Zone Abbr']).toBe('')

    const se = issueNoticePlaceholders(notice, pdf, manual, { zone: 'Quthbullapur' })
    expect(se['Zone']).toBe('Quthbullapur')
    expect(se['Zone Abbr']).toBe('QBZ')
    expect(se['Circle']).toBe('')
    expect(se['CNO']).toBe('')
  })
})

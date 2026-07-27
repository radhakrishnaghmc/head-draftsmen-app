import { describe, expect, it } from 'vitest'
import {
  integerToIndianWords,
  amountToWords,
  dateToWords,
  ordinal,
  parseDateParts
} from '../core/numberToWords'
import {
  deriveFields,
  workOrderPlaceholders,
  agreementPlaceholders,
  indianFinancialYear
} from '../core/workOrderAgreement'

describe('integerToIndianWords', () => {
  it('spells Indian-numbered figures', () => {
    expect(integerToIndianWords(0)).toBe('Zero')
    expect(integerToIndianWords(83)).toBe('Eighty Three')
    expect(integerToIndianWords(983)).toBe('Nine Hundred Eighty Three')
    expect(integerToIndianWords(1445983)).toBe('Fourteen Lakh Forty Five Thousand Nine Hundred Eighty Three')
    expect(integerToIndianWords(1631843)).toBe('Sixteen Lakh Thirty One Thousand Eight Hundred Forty Three')
    expect(integerToIndianWords(12345678)).toBe('One Crore Twenty Three Lakh Forty Five Thousand Six Hundred Seventy Eight')
  })
})

describe('amountToWords', () => {
  it('appends paise only when non-zero', () => {
    expect(amountToWords(1445983.17)).toBe(
      'Fourteen Lakh Forty Five Thousand Nine Hundred Eighty Three Rupees and Seventeen Paise Only'
    )
    expect(amountToWords(73658)).toBe('Seventy Three Thousand Six Hundred Fifty Eight Rupees Only')
  })
})

describe('date helpers', () => {
  it('ordinals', () => {
    expect(ordinal(1)).toBe('1st')
    expect(ordinal(2)).toBe('2nd')
    expect(ordinal(4)).toBe('4th')
    expect(ordinal(11)).toBe('11th')
    expect(ordinal(22)).toBe('22nd')
  })
  it('parses dd.mm.yyyy and ISO', () => {
    expect(parseDateParts('04.07.2026')).toEqual({ day: 4, month: 7, year: 2026 })
    expect(parseDateParts('2026-07-22')).toEqual({ day: 22, month: 7, year: 2026 })
    expect(parseDateParts('')).toBeNull()
    expect(parseDateParts('99.99.2026')).toBeNull()
  })
  it('spells dates in words', () => {
    expect(dateToWords('04.07.2026')).toBe('4th day of July 2026')
    expect(dateToWords('22.07.2026')).toBe('22nd day of July 2026')
    expect(dateToWords('')).toBe('')
  })
})

describe('indianFinancialYear', () => {
  it('splits on the 1 April boundary', () => {
    expect(indianFinancialYear(new Date(2026, 6, 27))).toBe('2026-27')
    expect(indianFinancialYear(new Date(2026, 2, 15))).toBe('2025-26')
  })
})

// A Works List row + L-1 selection matching the "Laying of Storm water line …"
// agreement sample (ECV 19,93,085; 27.45% less; contract 14,45,983.17).
const row: Record<string, string> = {
  Circle: 'Nizampet',
  CNO: '58',
  Zone: 'Quthbullapur',
  'Name of the work': 'Laying of Storm water line from Vinayaka Nagar to KNR Colony',
  'Amount of estimate': '25',
  ECV: '1993085',
  'Tender Percentage': '27.45',
  'Name of the Agency': 'Sri.Chiranjevi Alakuntla Works Contractor',
  'Address of the agency': 'Nizampet',
  'Phone number of the agency': '9000000000',
  Wincode: 'D058-26000001'
}

describe('deriveFields + placeholders', () => {
  const fields = deriveFields({}, { noticeDate: '04.07.2026' }, row)

  it('carries the L-1 date into both date fields', () => {
    expect(fields.workOrderDate).toBe('04.07.2026')
    expect(fields.agreementDate).toBe('04.07.2026')
    expect(fields.ecvRupees).toBe('1993085')
    expect(fields.contractRupees).toBe('1445983.17')
  })

  it('formats the Work Order the office way (plain 2-decimal Rs.)', () => {
    const wo = workOrderPlaceholders(fields)
    expect(wo['Estimate Amount']).toBe('Rs. 25.00 Lakhs')
    expect(wo['ECV']).toBe('Rs. 1993085.00')
    expect(wo['Tender Percentage']).toBe('(-) 27.45%-Less')
    expect(wo['TP']).toBe('27.45')
    expect(wo['Contract Amount']).toBe('Rs. 1445983.17')
    expect(wo['Circle']).toBe('Nizampet')
    expect(wo['Financialyear']).toMatch(/^\d{4}-\d{2}$/)
  })

  it('formats the Agreement the office way (grouped + words)', () => {
    const ag = agreementPlaceholders(fields)
    expect(ag['Estimate Amount']).toBe('25.00')
    expect(ag['ECV']).toBe('Rs.19,93,085.00')
    expect(ag['Tender percentage']).toBe('27.45')
    expect(ag['Contract value']).toBe('Rs.14,45,983.17')
    expect(ag['Agreement date in words']).toBe('4th day of July 2026')
    expect(ag['Contract value in rupees']).toContain('Rupees')
  })

  it('shows a nomination (0%) tender as "0%", not "(-) 0.00%-Less"', () => {
    const nomFields = deriveFields({}, {}, { ...row, ECV: '73658', 'Tender Percentage': '0' })
    const wo = workOrderPlaceholders(nomFields)
    expect(wo['Tender Percentage']).toBe('0%')
    expect(wo['Contract Amount']).toBe('Rs. 73658.00')
  })

  it('leaves amount fields blank when ECV is blank (never falls back to estimate)', () => {
    const blankEcv = deriveFields({}, {}, { ...row, ECV: '', 'Tender Percentage': '' })
    const wo = workOrderPlaceholders(blankEcv)
    expect(wo['ECV']).toBe('')
    expect(wo['Contract Amount']).toBe('')
    expect(wo['Estimate Amount']).toBe('Rs. 25.00 Lakhs')
  })
})

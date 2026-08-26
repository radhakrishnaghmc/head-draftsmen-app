import { describe, it, expect } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import PizZip from 'pizzip'
import {
  zoneAbbr,
  financialYearFromDate,
  formatIndianAmount,
  integerToIndianWords,
  amountInWords
} from '../core/loaSe'

describe('loaSe helpers', () => {
  it('zone abbreviation', () => {
    expect(zoneAbbr('Quthbullapur')).toBe('QBZ')
    expect(zoneAbbr('quthbullapur')).toBe('QBZ')
    expect(zoneAbbr('Kukatpally')).toBe('KPZ')
    expect(zoneAbbr('Serilingampally')).toBe('SLP')
    expect(zoneAbbr('')).toBe('')
  })

  it('financial year from a dd.mm.yyyy date', () => {
    expect(financialYearFromDate('03.07.2026')).toBe('2026-27')
    expect(financialYearFromDate('10.03.2026')).toBe('2025-26')
  })

  it('Indian grouping', () => {
    expect(formatIndianAmount(5267036, 2)).toBe('52,67,036.00')
    expect(formatIndianAmount(16491978, 2)).toBe('1,64,91,978.00')
    expect(formatIndianAmount(6597, 0)).toBe('6,597')
    expect(formatIndianAmount(165900, 0)).toBe('1,65,900')
  })

  it('EMD @1.5% and E-corpus @0.04% round to the letters', () => {
    // item 5: ECV 53,27,430 -> EMD 79,911, E-corpus 2,131
    expect(formatIndianAmount(Math.round(5327430 * 0.015), 0)).toBe('79,911')
    expect(formatIndianAmount(Math.round(5327430 * 0.0004), 0)).toBe('2,131')
    // item 7: ECV 1,10,59,985 -> EMD 1,65,900, ASD (26.3-25)% = 1,43,780
    expect(formatIndianAmount(Math.round(11059985 * 0.015), 0)).toBe('1,65,900')
    expect(formatIndianAmount(Math.round((11059985 * (26.3 - 25)) / 100), 0)).toBe('1,43,780')
  })

  it('integer to Indian words (CMC style: spaces, Lakhs plural)', () => {
    expect(integerToIndianWords(3976612)).toBe('Thirty Nine Lakhs Seventy Six Thousand Six Hundred Twelve')
    expect(integerToIndianWords(5392514)).toBe('Fifty Three Lakhs Ninety Two Thousand Five Hundred Fourteen')
  })

  it('amount in words with paise, as the letter wraps it', () => {
    expect(amountInWords(3976612.18)).toBe(
      'Thirty Nine Lakhs Seventy Six Thousand Six Hundred Twelve and Eighteen Paise'
    )
    expect(amountInWords(5392514.07)).toBe(
      'Fifty Three Lakhs Ninety Two Thousand Five Hundred Fourteen and Seven Paise'
    )
  })
})

describe('loa-se templates', () => {
  // Regression guard for a real request: the "Copy Submitted to the Chief
  // Engineer, CMC for favor of information please." line was removed from
  // both templates' signature blocks — re-check this stays gone after any
  // future re-export (bundled templates are known to regress that way).
  for (const name of ['loa-se-template.docx', 'loa-se-reserved-template.docx']) {
    it(`${name} no longer carries the "Copy Submitted to the Chief Engineer" line`, () => {
      const buf = readFileSync(resolve(__dirname, '..', 'resources', name))
      const xml = new PizZip(buf).file('word/document.xml')!.asText()
      expect(xml).not.toContain('Chief Engineer')
    })
  }
})

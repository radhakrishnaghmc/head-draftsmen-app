import { describe, it, expect } from 'vitest'
import fs from 'fs'
import path from 'path'
import { fillPlaceholdersInDocx } from '../core/docx-edit'
import { zoneAbbr, financialYearFromDate, formatIndianAmount, amountInWords } from '../core/loaSe'
import PizZip from 'pizzip'

// Figures lifted from a real MMC SE LOA sample (Intimation_works-26-27.doc,
// Malkajgiri Zone), verified against the app's EMD/ASD formulas before
// building the templates: reserved uses 2.5% EMD (recovered from RA bills)
// + ASD when tender% > 25, non-reserved uses 1.5% balance EMD + same ASD
// rule, with BG validity = period-of-completion months + 24.
function readXmlText(buf: Buffer): string {
  const xml = new PizZip(buf).file('word/document.xml')!.asText()
  return xml.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
}

describe('MMC LOA template end-to-end fill', () => {
  it('reserved: EMD recovered from RA bills at 2.5%, plus ASD when tender% > 25', () => {
    const ecv = 6088764, tenderPct = 27.89
    const emd = Math.round(ecv * 0.025) // 152219
    const asd = Math.round((ecv * (tenderPct - 25)) / 100) // 175965
    const values: Record<string, string> = {
      Corporation: 'MMC', Zone: 'Malkajgiri', 'Zone Abbr': zoneAbbr('Malkajgiri'),
      'Financial year': financialYearFromDate('23.04.2026'), 'LOA Date': '05.05.2026',
      'Agency Name': 'Sri. Kadari Ravinder Kumar', 'Address of the agency': 'Secunderabad-500026',
      'agency phone number': '9347585069',
      'Name of the work': 'Laying of CC Main road at Krishna nagar, Moula-ali Circle-04, MMC',
      'Reserved Tag': '(Reserved for SC category)', 'Item No': '1', 'Nit No': '04/SE(M)/Malkajgiri Zone/MMC/2026-27',
      'Nit Date': '23.04.2026', 'Price Bid opening date': '01-05-2026',
      ECV: formatIndianAmount(ecv, 2), 'Tender Percentage': String(tenderPct),
      'Contract Amount': formatIndianAmount(4390607.72, 2), 'Period Completion': '6 Months',
      'Contract In Words': amountInWords(4390607.72), 'Circle Line': 'Moula-Ali Circle-04',
      'MMC Reserved Emd Clause': `EMD Exempted. However, an amount of Rs.${formatIndianAmount(emd, 0)}/- towards 2½% EMD shall be recovered from respective RA bills as stipulated under G.O.Ms.No.59, I&CAD (Reforms) Department, Dt:21-05-2018. & Submitted ASD of Rs.${formatIndianAmount(asd, 0)}/- through online / Irrevocable Bank Guarantee from any nationalized / scheduled banks in favour of the Commissioner, MMC.`
    }
    const buf = fs.readFileSync(path.resolve(__dirname, '../resources/loa-se-mmc-reserved-template.docx'))
    const resolved = Object.keys(values).map((k) => ({ label: k, column: k, score: 1 }))
    const out = fillPlaceholdersInDocx(buf, resolved, values)
    const text = readXmlText(out)
    expect(text).not.toMatch(/\{\{/)
    expect(text).toContain('towards 2½% EMD shall be recovered from respective RA bills')
    expect(text).toContain('Rs.1,52,219/-')
    expect(text).toContain('Rs.1,75,965/-')
    expect(text).not.toContain('e-Corpus')
    expect(text).not.toContain('TSTS')
  })

  it('non-reserved: 1.5% balance EMD + ASD, BG validity = period + 24', () => {
    const ecv = 4518694, tenderPct = 33.79, period = 3
    const emd = Math.round(ecv * 0.015) // 67780
    const asd = Math.round((ecv * (tenderPct - 25)) / 100) // 397193
    const values: Record<string, string> = {
      Corporation: 'MMC', Zone: 'Malkajgiri', 'Zone Abbr': zoneAbbr('Malkajgiri'),
      'Financial year': financialYearFromDate('16.04.2026'), 'LOA Date': '05.05.2026',
      'Agency Name': 'Sri. Shyam Kumar Miryala', 'Address of the agency': 'Medchal - Malkajgiri.',
      'agency phone number': '9885595414',
      'Name of the work': 'Laying of CC road at Bhanu Enclave, Keesara Circle-01, MMC',
      'Reserved Tag': '', 'Item No': '7', 'Nit No': '03/SE(M)/Malkajgiri Zone/MMC/2026-27',
      'Nit Date': '16.04.2026', 'Price Bid opening date': '24-04-2026',
      ECV: formatIndianAmount(ecv, 2), 'Tender Percentage': String(tenderPct),
      'Contract Amount': formatIndianAmount(2991827.30, 2), 'Period Completion': `${period} Months`,
      'Period Number': String(period), 'Contract In Words': amountInWords(2991827.30),
      'Circle Line': 'Keesara Circle-01',
      'MMC Emd Clause': `of Rs.${formatIndianAmount(emd, 0)}/- & ASD of Rs.${formatIndianAmount(asd, 0)}/-`,
      'BG Validity Months': String(period + 24)
    }
    const buf = fs.readFileSync(path.resolve(__dirname, '../resources/loa-se-mmc-template.docx'))
    const resolved = Object.keys(values).map((k) => ({ label: k, column: k, score: 1 }))
    const out = fillPlaceholdersInDocx(buf, resolved, values)
    const text = readXmlText(out)
    expect(text).not.toMatch(/\{\{/)
    expect(text).toContain('Balance 1½% EMD of Rs.67,780/- & ASD of Rs.3,97,193/-')
    expect(text).toContain('27 Months (i.e., 3+24=27)')
    expect(text).toContain('in addition to 1% EMD already paid')
    expect(text).toContain('within 3 days')
    expect(text).not.toContain('e-Corpus')
    expect(text).not.toContain('TSTS')
  })
})

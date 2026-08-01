import { describe, it, expect } from 'vitest'
import fs from 'fs'; import path from 'path'
import { fillPlaceholdersInDocx } from '../core/docx-edit'
import { zoneAbbr, financialYearFromDate, formatIndianAmount, amountInWords } from '../core/loaSe'
import PizZip from 'pizzip'

// Mirror of resolveLoaValue's math to end-to-end fill the real template and
// assert the rendered text (docx-edit + template) is right for a non-reserved work.
const office = { corporation: 'CMC', zone: 'Quthbullapur' }
const ecv = 5327430, tenderPct = 27, contract = 3889023.9
const emd = Math.round(ecv * 0.015)                        // 79911
const asd = Math.round((ecv * (tenderPct - 25)) / 100)      // 106549
const values: Record<string, string> = {
  'Zone Abbr': zoneAbbr(office.zone), Zone: office.zone, 'Financial year': financialYearFromDate('03.07.2026'),
  'LOA Date': '14.07.2026.', 'Agency Name': 'Sri N.Vengal Rao', 'Address of the agency': 'H.No:4-35-70/1,\nHyderabad.',
  'agency phone number': '9951009833', 'Name of the work': 'Laying of CC roads, Gajularamaram Circle-57, CMC',
  'Reserved Tag': '', 'Item No': '5', 'Nit No': '12/SE/QBZ/CMC/2026-27', 'Nit Date': '03.07.2026',
  'Tender ID': '714894', 'Price Bid opening date': '14.07.2026', 'Admin Sanction Value': formatIndianAmount(6800000, 2),
  ECV: formatIndianAmount(ecv, 2), 'Tender Percentage': String(tenderPct), 'Contract Amount': formatIndianAmount(contract, 2),
  'Period of Completion': '03 Months', 'Contract In Words': amountInWords(contract),
  'EMD Clause': `Rs. ${formatIndianAmount(emd, 0)}/- & ASD amount of Rs.${formatIndianAmount(asd, 0)}/-`,
  'E-Corpus': formatIndianAmount(Math.round(ecv * 0.0004), 0), 'Circle Line': 'Gajularamaram Circle-57'
}

describe('LOA template end-to-end fill', () => {
  it('fills the non-reserved template with no leftover placeholders and correct figures', () => {
    const buf = fs.readFileSync(path.resolve(__dirname, '../resources/loa-se-template.docx'))
    const resolved = Object.keys(values).map((k) => ({ label: k, column: k, score: 1 }))
    const out = fillPlaceholdersInDocx(buf, resolved, values)
    const xml = new PizZip(out).file('word/document.xml')!.asText()
    const text = xml.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    expect(text).not.toMatch(/\{\{/)                                   // every placeholder filled
    expect(text).toContain('Rs. 79,911/- & ASD amount of Rs.1,06,549/-')
    expect(text).toContain('Rs.52,67,430.00'.replace('52,67,430', '53,27,430'))
    expect(text).toContain('Thirty Eight Lakhs Eighty Nine Thousand Twenty Three and Ninety Paise')
    expect(text).toContain('Gajularamaram Circle-57')
  })

  it('reserved template has no EMD clause placeholder', () => {
    const buf = fs.readFileSync(path.resolve(__dirname, '../resources/loa-se-reserved-template.docx'))
    const xml = new PizZip(buf).file('word/document.xml')!.asText().replace(/<[^>]+>/g, '')
    expect(xml).not.toContain('{{EMD Clause}}')
    expect(xml).toContain('e-Corpus Fund')
  })
})

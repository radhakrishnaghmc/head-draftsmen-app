import { describe, expect, it } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import { fillBidDocument, lakhsToRupees } from '../core/bidDocument'
import { listParagraphs } from '../core/docx-edit'

const TEMPLATE_PATH = resolve(__dirname, '../resources/bid-document-template.docx')

describe('lakhsToRupees', () => {
  it('converts a Lakhs figure to a plain rupee number', () => {
    expect(lakhsToRupees('45')).toBe(4500000)
    expect(lakhsToRupees('3.2')).toBe(320000)
    expect(lakhsToRupees('')).toBe(0)
  })
})

describe('fillBidDocument', () => {
  const buffer = readFileSync(TEMPLATE_PATH)

  it('fills the work-specific fields, converting the estimate from Lakhs to Indian-formatted rupees', () => {
    const out = fillBidDocument(buffer, {
      nitNo: '16/DB/EE/Gajularamaram Circle-57/CMC/2026-27',
      dated: '21.07.2026',
      downloadStartDate: '16.07.2026',
      downloadEndDate: '23.07.2026',
      work: {
        serial: 1,
        name: 'Test civil work',
        amount: '45',
        zone: 'Cyberabad',
        circle: 'Gajularamaram Circle-57',
        completionPeriod: '3'
      }
    })

    const paragraphs = listParagraphs(out)
    const joined = paragraphs.join('\n')

    expect(joined).toContain('Test civil work')
    expect(joined).not.toContain('{{Name of the work}}')
    // 45 Lakhs -> 4500000. The cover page has no "Rs:" of its own, so it gets the full "Rs 45,00,000/-".
    expect(joined).toContain('Rs 45,00,000/-')
    expect(joined).not.toContain('{{Estimate Amount}}')
    // No ECV supplied -> {{ECV}} and {{EMD 1%}} are left blank (just the template's own "Rs:"
    // label with nothing after it), never computed from the 45-Lakh estimate instead — ECV and
    // the estimate are distinct figures.
    expect(paragraphs.filter((p) => p.trim() === 'Rs:')).toHaveLength(2)
    expect(joined).not.toContain('{{ECV}}')
    expect(joined).not.toContain('{{EMD 1%}}')
    expect(joined).toContain('16/DB/EE/Gajularamaram Circle-57/CMC/2026-27')
    expect(joined).not.toContain('{{Nit no.}}')
    expect(joined).toContain('16.07.2026 at 2:00 P.M')
    expect(joined).toContain('23.07.2026 at 2:00 P.M')
    expect(joined).toContain('23.07.2026 at 2:30 P.M')
    expect(joined).toContain('Cyberabad')
    expect(joined).toContain('Gajularamaram Circle-57')
    expect(joined).not.toContain('{{Zone}}')
    expect(joined).not.toContain('{{Circle}}')
  })

  it('fills {{ECV}} and computes EMD @ 1% from it, distinct from the plain estimate amount', () => {
    const out = fillBidDocument(buffer, {
      nitNo: '1/EE/2026',
      dated: '01.01.2026',
      downloadStartDate: '02.01.2026',
      downloadEndDate: '09.01.2026',
      work: {
        serial: 2,
        name: 'Another work',
        amount: '10',
        ecv: '1200000', // ECV is stored in rupees (Amount of estimate is in Lakhs)
        completionPeriod: '2'
      }
    })
    const paragraphs = listParagraphs(out)
    // Estimate Amount (cover page) stays at the 10-Lakh estimate.
    expect(paragraphs.join('\n')).toContain('Rs 10,00,000/-')
    // ECV 12,00,000 rupees, distinct from the estimate.
    const ecvLine = paragraphs.find((p) => p.startsWith('Rs:') && p.includes('12,00,000/-'))
    expect(ecvLine).toBeDefined()
    // EMD @ 1% of ECV = 12,000 (not 1% of the 10-Lakh estimate, which would be 10,000).
    const emdLine = paragraphs.find((p) => p.startsWith('Rs:') && p.includes('12,000/-'))
    expect(emdLine).toBeDefined()
  })

  it('fills the work name into the Conditions of Contract covering-letter blank', () => {
    const out = fillBidDocument(buffer, {
      nitNo: '22/DB/EE/Nizampet Circle-58/CMC/2026-27',
      dated: '10.08.2026',
      downloadStartDate: '18.08.2026',
      downloadEndDate: '25.08.2026',
      work: { serial: 1, name: 'Providing temporary lighting at Nizampet', amount: '5' }
    })
    const para = listParagraphs(out).find((p) =>
      p.includes('hereby tender and if this tender be accepted')
    )
    expect(para).toBeDefined()
    // The name lands right in the blank after "viz" (previously an empty underscore line).
    expect(para).toContain('viz“Providing temporary lighting at Nizampet.”')
  })
})

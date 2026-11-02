import { describe, it, expect } from 'vitest'
import { readFileSync } from 'fs'
import { fillPlaceholdersInAllParts, extractAllText, findPlaceholdersInAllParts } from '../core/docx-edit'

function fill(file: string, values: Record<string, string>): string {
  const buf = readFileSync(`resources/${file}`)
  const labels = findPlaceholdersInAllParts(buf)
  const resolved = labels.map((label) => ({ label, column: label, score: 1 }))
  const row: Record<string, string> = Object.fromEntries(labels.map((l) => [l, values[l] ?? '']))
  return extractAllText(fillPlaceholdersInAllParts(buf, resolved, row))
}

describe('Eligibility Criteria — SC/ST reserved variant', () => {
  const values = { 'ECV Lakhs': '40.07', 'Reserved Tag': '(Reserved for SC only)', 'Name of the work': 'Laying CC road' }
  const reserved = fill('eligibility-criteria-reserved-template.docx', values)
  const open = fill('eligibility-criteria-template.docx', values)

  it('exempts EMD instead of printing an amount', () => {
    expect(reserved).toMatch(/EXEMPTED/)
    expect(reserved).toContain('G.O.Ms.No.59')
    expect(reserved).not.toContain('Online payment/BG')
  })

  it('adds the caste-certificate and co-op audit clauses, and the G.O. 66 note', () => {
    expect(reserved).toContain('12. Caste certificate')
    expect(reserved).toMatch(/13\. All SC\/ST\/Waddera\/Uppera\/Sagar/)
    expect(reserved).toContain('66, Irrigation & CAD (Reforms) Department, Dated: 20.04.2015.')
  })

  it('carries the reservation tag after the work name and the class-wise solvency rule', () => {
    expect(reserved).toContain('Laying CC road (Reserved for SC only)')
    expect(reserved).toContain('solvency is not needed')
  })

  it('follows the reserved-work note format: circle/tag in Sub, TS authority, five-item equipment, two-signatory block', () => {
    const r = fill('eligibility-criteria-reserved-template.docx', {
      ...values,
      'EE Circle': 'Alwal Circle-02',
      'TS Authority': 'SE (M), Malkajgiri Zone, MMC',
      'TS Ref No': '40/SE(M)/MKJZone/MMC/2026-27',
      'AS E-Office Ref': ', vide e-office file no.403845',
      'Equipment 4': 'Air compressor',
      'Equipment 5': 'Water tanker'
    })
    expect(r).toContain('Alwal Circle-02 – Laying CC road (Reserved for SC only) – Submitted for Eligibility Criteria – Approval Requested – Reg.')
    expect(r).toContain('vide e-office file no.403845')
    expect(r).toContain('Technical Sanction accorded by the SE (M), Malkajgiri Zone, MMC vide TS.No.40/SE(M)/MKJZone/MMC/2026-27')
    expect(r).toContain('Air compressor')
    expect(r).toContain('AE/AEE/DyEE')
    expect(r).not.toContain('Req-Reg')
    expect(r).not.toContain('owner ship')
  })

  it('leaves the open-category template unchanged', () => {
    expect(open).not.toContain('Caste certificate')
    expect(open).not.toContain('EXEMPTED')
  })
})

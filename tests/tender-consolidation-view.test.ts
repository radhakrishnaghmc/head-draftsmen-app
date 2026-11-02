import { describe, it, expect } from 'vitest'
import { execFileSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { parseTenderConsolidationView, parseConditionsByColumn, completionYears, type PositionedItem } from '../core/tenderConsolidationView'

const LINES = [
  'Tender ID 742254',
  'Period of Completion/ Delivery Period (In',
  'Months)',
  '3',
  'Required Tender Documents Details',
  'S.No Documents Required from Contractor Stage Optional',
  '1 Registration Certification COMMON Mandatory',
  '2',
  '2.Class-V of GHMC/CMC registration is only',
  'considered, Other Departments Registration Class-III',
  '& above only are eligible',
  'COMMON Mandatory',
  '3 3.EMD COMMON Mandatory4 4.Pan Card to be uploaded COMMON Mandatory',
  '13',
  'If the Agency/contractor quotes more than 25 percent',
  'less over ECV.',
  'COMMON Mandatory',
  'Geographical Particulars'
]

describe('parseTenderConsolidationView', () => {
  it('reads tender id, period and the required-documents conditions', () => {
    const r = parseTenderConsolidationView(LINES)
    expect(r.tenderId).toBe('742254')
    expect(r.completionMonths).toBe(3)
    expect(r.conditions).toEqual([
      'Registration Certification',
      'Class-V of GHMC/CMC registration is only considered, Other Departments Registration Class-III & above only are eligible',
      'EMD',
      'Pan Card to be uploaded',
      'If the Agency/contractor quotes more than 25 percent less over ECV.'
    ])
    expect(completionYears(r.completionMonths)).toBe(0.25)
  })

  const real = '/Users/radhakrishnapodugu/Library/CloudStorage/GoogleDrive-radhakrishnaghmc@gmail.com/Other computers/My Computer/Keerthi Kapa NAC ENGG/DUNDIGAL CIRCLE-59 CMC/TenderConsolidationView_742254.pdf'
  it.skipIf(!existsSync(real))('parses the real 742254 PDF into 13 conditions', () => {
    const txt = execFileSync('pdftotext', ['-raw', real, '-'], { encoding: 'utf8' })
    const r = parseTenderConsolidationView(txt.split('\n'))
    expect(r.tenderId).toBe('742254')
    expect(r.completionMonths).toBe(3)
    expect(r.conditions).toHaveLength(13)
    expect(r.conditions[7]).toMatch(/^Bid capacity \(3AN-B\)/)
  })

  // pdf.js isn't usable under node here, so positions come from pdftotext -bbox (word level, like pdf.js items).
  const itemsOf = (pdf: string): PositionedItem[] => {
    const html = execFileSync('pdftotext', ['-bbox', pdf, '-'], { encoding: 'utf8' })
    const out: PositionedItem[] = []
    html.split('<page ').slice(1).forEach((pg, n) => {
      for (const m of pg.matchAll(/<word xMin="([\d.]+)" yMin="([\d.]+)"[^>]*>([^<]*)<\/word>/g)) {
        out.push({ text: m[3].replace(/&amp;/g, '&'), x: Number(m[1]), y: n * 100000 + Number(m[2]) })
      }
    })
    return out
  }
  const dir = '/Users/radhakrishnapodugu/Library/CloudStorage/GoogleDrive-radhakrishnaghmc@gmail.com/Other computers/My Computer/Keerthi Kapa NAC ENGG/DUNDIGAL CIRCLE-59 CMC/'
  it.skipIf(!existsSync(real))('reads wrapped, vertically-centred rows by column (738797)', () => {
    const c = parseConditionsByColumn(itemsOf(dir + 'TenderConsolidationView_738797.pdf'))
    expect(c).toHaveLength(13)
    expect(c[2]).toBe('EMD')
    for (const x of c) expect(x).not.toMatch(/COMMON|Mandatory/)
  })
})

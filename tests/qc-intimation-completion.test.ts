import { describe, it, expect } from 'vitest'
import fs from 'fs'
import { resolve } from 'path'
import { findPlaceholdersInAllParts, fillPlaceholdersInAllParts, extractAllText } from '../core/docx-edit'
import { qcIntimationCompletionPlaceholders, qcPartyToBlock, type WorkOrderAgreementFields } from '../core/workOrderAgreement'

const template = () => fs.readFileSync(resolve(__dirname, '../resources/qc-intimation-completion-template.docx'))

describe('qcPartyToBlock', () => {
  it('prints the address block once when it already names the agency', () => {
    const out = qcPartyToBlock({ name: 'M/s Coherent Infrastructure,Solution Pvt Ltd,', address: 'M/s Coherent Infrastructure,\nSolution Pvt Ltd,\n10-3-76 ,Mehadipatnam.\n' })
    expect(out).toBe('M/s Coherent Infrastructure,\nSolution Pvt Ltd,\n10-3-76 ,Mehadipatnam.')
  })
  it('puts the name on top when the address block does not mention it', () => {
    expect(qcPartyToBlock({ name: 'ABC Labs', address: 'Plot 4,\nHyderabad' })).toBe('ABC Labs\nPlot 4,\nHyderabad')
  })
  it('is empty when nothing is entered', () => {
    expect(qcPartyToBlock(undefined)).toBe('')
  })
})

describe('QC Intimation/Completion template', () => {
  it('only uses placeholders the app fills', () => {
    const known = new Set(Object.keys(qcIntimationCompletionPlaceholders({ corporationFullName: '', completionMonths: '' } as WorkOrderAgreementFields)).map((k) => k.toLowerCase()))
    for (const label of findPlaceholdersInAllParts(template())) expect(known.has(label.toLowerCase()), label).toBe(true)
  })

  it('fills all five letters, leaving no placeholder or stale sample text behind', () => {
    const f = {
      corporationFullName: 'Greater Hyderabad Municipal Corporation', corporation: 'GHMC', zone: 'Kukatpally', circle: 'Moosapet', cno: '23', financialYear: '2026-27',
      nameOfWork: 'Laying of BT road at Test Colony', estimateLakhs: '25', ecvRupees: '2000000', contractRupees: '1800000.5', wincode: 'W1', agencyName: 'M/s Test Agency', completionMonths: '4',
      qcThirdTo: 'Professor & Head,\nCollege X', qcFourthTo: 'M/s Lab Y,\nAddress Y'
    } as WorkOrderAgreementFields
    const buf = template()
    const out = fillPlaceholdersInAllParts(buf, findPlaceholdersInAllParts(buf).map((l) => ({ label: l, column: l, score: 1 })), qcIntimationCompletionPlaceholders(f))
    const text = extractAllText(out)
    expect(text).not.toMatch(/\{\{/)
    for (const stale of ['Nizampet', 'Quthbullapur', 'CYBERABAD', 'Siri Enclave', 'Giridhar', 'Coherent', 'VNR', 'Cir-58']) expect(text, stale).not.toContain(stale)
    expect(text.match(/INTIMATION\s+LETTER/g)).toHaveLength(2)
    expect(text.match(/COMPLETION\s+LETTER/g)).toHaveLength(3)
    expect(text).toContain('GREATER HYDERABAD MUNICIPAL CORPORATION')
    expect(text).toContain('Rs 25.00 Lakhs')
    expect(text).toContain('Rs 18,00,000.50')
    // Intimation letters are from/signed by the Deputy EE, Completion letters by the EE.
    expect(text.match(/Deputy\s+Executive Engineer/g)).toHaveLength(4)
    expect(text.match(/Executive Engineer/g)!.length - 4).toBeGreaterThanOrEqual(6)
    expect(text).toContain('College X')
    expect(text).toContain('Lab Y')
  })
})

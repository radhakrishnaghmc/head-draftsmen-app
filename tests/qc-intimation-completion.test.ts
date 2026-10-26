import { describe, it, expect } from 'vitest'
import fs from 'fs'
import { resolve } from 'path'
import { findPlaceholdersInAllParts, fillPlaceholdersInAllParts, extractAllText } from '../core/docx-edit'
import { qcIntimationCompletionPlaceholders, qcPartyToBlock, type WorkOrderAgreementFields } from '../core/workOrderAgreement'

const template = () => fs.readFileSync(resolve(__dirname, '../legacy-templates/qc-intimation-completion-template.docx'))

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

describe('QC letters quote the Work Order in Ref', () => {
  const f = { corporationFullName: '', nameOfWork: 'W', estimateLakhs: '25', ecvRupees: '2000000', contractRupees: '1800000', wincode: 'W1', agencyName: 'A', completionMonths: '4', circle: 'Moosapet', cno: '23', zone: 'Kukatpally', corporation: 'GHMC', financialYear: '2026-27', agreementNo: '37', agreementDate: '12/09/2026', workOrderDate: '05/09/2026' } as WorkOrderAgreementFields
  for (const v of ['cmc', 'ghmc', 'mmc']) {
    it(`${v}: all five letters carry the work order number and date`, () => {
      const buf = fs.readFileSync(resolve(__dirname, `../resources/qc-intimation-completion-template-${v}.docx`))
      const text = extractAllText(fillPlaceholdersInAllParts(buf, findPlaceholdersInAllParts(buf).map((l) => ({ label: l, column: l, score: 1 })), qcIntimationCompletionPlaceholders(f)))
      expect(text.match(/Work Order No\s+\/GHMC\/KPZ\/Moosapet\/CIR-23\/EE\/2026-27 Dt: 05\/09\/2026/g)).toHaveLength(5)
      expect(text).not.toContain('Agreement No')
    })
  }
  it('the 4th-party table carries the work order number and the typed AE name & phone', () => {
    const buf = fs.readFileSync(resolve(__dirname, '../resources/qc-intimation-completion-template-cmc.docx'))
    const values = qcIntimationCompletionPlaceholders({ ...f, aeNamePhone: ' K. Ramesh, 9876543210 ' })
    const text = extractAllText(fillPlaceholdersInAllParts(buf, findPlaceholdersInAllParts(buf).map((l) => ({ label: l, column: l, score: 1 })), values))
    expect(text).toContain('K. Ramesh, 9876543210')
    // the W.O No cell repeats the Ref line's number and date, so 5 Ref lines + 1 table cell
    expect(text.match(/\/GHMC\/KPZ\/Moosapet\/CIR-23\/EE\/2026-27 Dt: 05\/09\/2026/g)).toHaveLength(6)
    expect(qcIntimationCompletionPlaceholders(f)['AE Name Phone']).toBe('')
  })
  it('the serial is left as a gap to write in, and a blank date prints no Dt', () => {
    expect(qcIntimationCompletionPlaceholders({ ...f, workOrderDate: '' })['Work Order No and Date']).toMatch(/^\u00a0{2}\/GHMC\/.*\/EE\/2026-27$/)
  })
})

describe('QC Intimation/Completion template', () => {
  it('only uses placeholders the app fills', () => {
    const known = new Set(Object.keys(qcIntimationCompletionPlaceholders({ corporationFullName: '', completionMonths: '', agreementDate: '' } as WorkOrderAgreementFields)).map((k) => k.toLowerCase()))
    for (const label of findPlaceholdersInAllParts(template())) expect(known.has(label.toLowerCase()), label).toBe(true)
  })

  it('fills all five letters, leaving no placeholder or stale sample text behind', () => {
    const f = {
      corporationFullName: 'Greater Hyderabad Municipal Corporation', corporation: 'GHMC', zone: 'Kukatpally', circle: 'Moosapet', cno: '23', financialYear: '2026-27',
      nameOfWork: 'Laying of BT road at Test Colony', estimateLakhs: '25', ecvRupees: '2000000', contractRupees: '1800000.5', wincode: 'W1', agencyName: 'M/s Test Agency', completionMonths: '4', agreementNo: '37', agreementDate: '12/09/2026',
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

import { describe, expect, it } from 'vitest'
import fs from 'fs'
import path from 'path'
import { workBillNotePlaceholders, type WorkOrderAgreementFields } from '../core/workOrderAgreement'
import { extractAllText, fillPlaceholdersInAllParts, findPlaceholdersInAllParts } from '../core/docx-edit'

const f = {
  circle: 'Quthbullapur', cno: '25', corporation: 'GHMC', wincode: 'D25-24038282', financialYear: '2025-26',
  nameOfWork: 'Laying CC Road at Angadipet', agencyName: 'G.Ganesh', estimateLakhs: '25', ecvRupees: '2000611',
  tenderPercent: '13.99', contractRupees: '1720725.52', agreementDate: '04.04.2025', completionMonths: '2'
} as WorkOrderAgreementFields

describe('workBillNotePlaceholders', () => {
  it('derives the work-level items and leaves bill-specific ones blank', () => {
    const v = workBillNotePlaceholders(f)
    expect(v.Wincode).toBe('D25-24038282')
    expect(v['Tender Percentage']).toBe('-13.99% less')
    expect(v.ECV).toBe('Rs.20,00,611.00')
    expect(v['Contract Amount Plain']).toBe('17,20,725.52')
    expect(v['Period of Contract']).toBe('2 months')
    expect(v['Value of Work Done']).toBe('')
  })
  it('never fills ECV from the estimate', () => {
    expect(workBillNotePlaceholders({ ...f, ecvRupees: '' }).ECV).toBe('')
  })
  it('every template placeholder has a value entry and fills away', () => {
    const buf = fs.readFileSync(path.join(__dirname, '../resources/work-bill-note-template.docx'))
    const labels = findPlaceholdersInAllParts(buf)
    const v = workBillNotePlaceholders(f)
    expect(labels.filter((l) => !(l in v))).toEqual([])
    const out = fillPlaceholdersInAllParts(buf, labels.map((label) => ({ label, column: label, score: 1 })), v)
    expect(findPlaceholdersInAllParts(out)).toEqual([])
    const text = extractAllText(out)
    expect(text).toContain('D25-24038282')
    expect(text).toContain('EE-25')
  })
})

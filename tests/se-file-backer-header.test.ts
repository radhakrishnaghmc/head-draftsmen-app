import { describe, expect, it } from 'vitest'
import * as fs from 'fs'
import PizZip from 'pizzip'
import { tightenSeFileBackerHeader } from '../core/docMargins'
import { deriveFields, fileBackerPlaceholders, type WorkOrderAgreementFields } from '../core/workOrderAgreement'

const fieldsWith = (o: Partial<WorkOrderAgreementFields>): WorkOrderAgreementFields => ({
  ...deriveFields({}, {}, {}),
  ...o
})

describe('SE File Backer header', () => {
  it('names the SE office when the office has a zone but no circle', () => {
    const f = fieldsWith({
      circle: '',
      cno: '',
      zone: 'Quthbullapur',
      corporation: 'CMC',
      corporationFullName: 'Cyberabad Municipal Corporation'
    })
    expect(fileBackerPlaceholders(f)['Circle Header']).toBe('Office of Superintending Engineer, Quthbullapur Zone, CMC')
  })

  it('keeps the EE circle line', () => {
    const f = fieldsWith({ circle: 'Gajularamaram', cno: '57', zone: '', corporation: 'CMC' })
    expect(fileBackerPlaceholders(f)['Circle Header']).toBe('GAJULARAMARAM CIRCLE-57')
  })

  it('shrinks the spacer paragraphs and the office line in every corporation template', () => {
    for (const c of ['cmc', 'ghmc', 'mmc']) {
      const b64 = fs.readFileSync(`resources/file-backer-template-${c}.docx`).toString('base64')
      const out = new PizZip(tightenSeFileBackerHeader(b64), { base64: true }).file('word/document.xml')!.asText()
      expect(out).not.toContain('<w:sz w:val="48"/>')
      expect(out).toMatch(/<w:sz w:val="28"\/><w:szCs w:val="28"\/><\/w:rPr><w:t>\{\{Circle Header\}\}/)
      expect(out).toContain('<w:b/><w:sz w:val="8"/><w:szCs w:val="8"/></w:rPr></w:pPr></w:p>')
    }
  })
})

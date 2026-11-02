import { describe, expect, it } from 'vitest'
import * as fs from 'fs'
import PizZip from 'pizzip'
import { removeDocParagraph } from '../core/docMargins'

const xmlOf = (b64: string) => new PizZip(b64, { base64: true }).file('word/document.xml')!.asText()

describe('removeDocParagraph', () => {
  const b64 = fs.readFileSync('resources/zonal-work-order-template-cmc.docx').toString('base64')

  it('drops only the Chief Engineer copy line from the SE Work Order', () => {
    const out = xmlOf(removeDocParagraph(b64, 'The Chief Engineer, {{Corp}} for favour of kind information'))
    expect(out).not.toContain('The Chief Engineer, {{Corp}}')
    expect(out).toContain('The Zonal Commissioner, {{Zone}} Zone')
    expect(out).toContain('Copy submitted to')
  })

  it('leaves the document alone when the text is absent', () => {
    expect(removeDocParagraph(b64, 'no such paragraph')).toBe(b64)
  })
})

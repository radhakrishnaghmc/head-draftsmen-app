import { describe, it, expect } from 'vitest'
import { readFileSync } from 'fs'
import PizZip from 'pizzip'
import { applyPageMargins } from '../core/docMargins'

const pgMar = (b64: string): string => new PizZip(b64, { base64: true }).file('word/document.xml')!.asText().match(/<w:pgMar[^>]*>/)![0]

describe('applyPageMargins on A4-forced Letter templates', () => {
  const fb = readFileSync('resources/file-backer-template-cmc.docx').toString('base64')
  it('narrows the side margins so the 9630-twip File Backer tables sit inside the page, centred', () => {
    const mar = pgMar(applyPageMargins(fb, null, true))
    expect(mar).toContain('w:left="1138"')
    expect(mar).toContain('w:right="1138"')
  })
  it('leaves hand-set margins exactly as given', () => {
    const mar = pgMar(applyPageMargins(fb, { top: 10, bottom: 10, left: 30, right: 30 }, true))
    expect(mar).toContain('w:left="1701"')
  })
  it('leaves a template whose tables already fit untouched', () => {
    const ag = readFileSync('resources/agreement-template.docx').toString('base64')
    expect(pgMar(applyPageMargins(ag, null, true))).toContain('w:left="1440"')
  })
})

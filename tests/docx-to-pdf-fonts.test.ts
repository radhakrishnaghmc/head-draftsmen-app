import { describe, it, expect } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import PizZip from 'pizzip'
import { substituteFontsForLibreOffice } from '../core/docxToPdf'

const FILE_BACKER = resolve(__dirname, '../resources/file-backer-template-header2.docx')
const INTIMATION = resolve(__dirname, '../resources/intimation-template.docx')

function part(buffer: Buffer, name: string): string {
  return new PizZip(buffer).file(name)?.asText() ?? ''
}

describe('substituteFontsForLibreOffice', () => {
  it('renames the Gautami Telugu caption font to Noto Sans Telugu, only in the given buffer', () => {
    const buf = readFileSync(FILE_BACKER)
    const before = part(buf, 'word/document.xml')
    expect(before).toContain('w:cs="Gautami"')

    const out = substituteFontsForLibreOffice(buf)
    const after = part(out, 'word/document.xml')
    expect(after).not.toContain('Gautami')
    expect(after).toContain('w:cs="Noto Sans Telugu"')

    // the original buffer (what a caller might separately save as the real
    // .docx) is left untouched — still Word-2007-safe with Gautami declared
    expect(part(buf, 'word/document.xml')).toContain('w:cs="Gautami"')
  })

  it('renames every Gautami font attribute (ascii/hAnsi/cs/eastAsia), not just w:cs', () => {
    const buf = readFileSync(INTIMATION)
    const before = part(buf, 'word/document.xml')
    expect(before).toContain('w:ascii="Gautami"')
    expect(before).toContain('w:hAnsi="Gautami"')

    const after = part(substituteFontsForLibreOffice(buf), 'word/document.xml')
    expect(after).not.toContain('Gautami')
    expect(after).toContain('w:ascii="Noto Sans Telugu"')
    expect(after).toContain('w:hAnsi="Noto Sans Telugu"')
  })

  it('is a no-op (returns the same buffer) when the template has no Gautami font at all', () => {
    const buf = Buffer.from(
      new PizZip()
        .file(
          '[Content_Types].xml',
          '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
        )
        .file(
          'word/document.xml',
          '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>'
        )
        .generate({ type: 'nodebuffer' })
    )
    const out = substituteFontsForLibreOffice(buf)
    expect(out).toBe(buf)
  })
})

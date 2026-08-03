import { describe, it, expect } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import PizZip from 'pizzip'
import { fillPlaceholdersInDocx, findPlaceholdersInDocx } from '../core/docx-edit'
import { sanitizeDocxForWord2007 } from '../core/word2007Compat'
import type { PlaceholderMatch } from '../core/createDocument'

const TENDER = resolve(__dirname, '../resources/civil-tender-template.docx')

function part(buffer: Buffer, name: string): string {
  return new PizZip(buffer).file(name)?.asText() ?? ''
}

describe('sanitizeDocxForWord2007', () => {
  it('rewrites bidirectional table-border sides to physical left/right', () => {
    const buf = readFileSync(TENDER)
    const before = part(buf, 'word/document.xml')
    // Precondition: the template really does carry the offending elements.
    expect(before).toContain('<w:start ')
    expect(before).toContain('<w:end ')

    const out = sanitizeDocxForWord2007(buf)
    const after = part(out, 'word/document.xml')

    // No bidirectional border side survives...
    expect(after).not.toContain('<w:start ')
    expect(after).not.toContain('<w:end ')
    // ...they became the physical sides Word 2007 understands.
    expect(after).toContain('<w:left ')
    expect(after).toContain('<w:right ')
    // Nothing dropped: every start became a left and every end a right, added
    // on top of whatever physical sides the template already had.
    const n = (s: string, re: RegExp) => (s.match(re) ?? []).length
    expect(n(after, /<w:left /g)).toBe(n(before, /<w:left /g) + n(before, /<w:start /g))
    expect(n(after, /<w:right /g)).toBe(n(before, /<w:right /g) + n(before, /<w:end /g))
  })

  it('rewrites bidirectional justification values to left/right', () => {
    const buf = readFileSync(TENDER)
    const before = part(buf, 'word/document.xml')
    // Precondition: the template uses the bidi justification values.
    expect(before).toContain('<w:jc w:val="start"')

    const after = part(sanitizeDocxForWord2007(buf), 'word/document.xml')
    expect(after).not.toContain('w:val="start"')
    expect(after).not.toContain('w:val="end"')
    // "both"/"center" justification is valid in Word 2007 — must survive.
    expect(after).toContain('<w:jc w:val="both"')
  })

  it('fixes list-level justification but preserves list start values', () => {
    const buf = readFileSync(TENDER)
    const numBefore = part(buf, 'word/numbering.xml')
    // Precondition: numbering carries both bidi <w:lvlJc w:val="start"> AND the
    // legitimate list start value <w:start w:val="1"/>.
    expect(numBefore).toContain('<w:lvlJc w:val="start"')
    expect(numBefore).toMatch(/<w:start w:val="\d/)
    const startVals = (numBefore.match(/<w:start w:val="\d+"\s*\/>/g) ?? []).length

    const numAfter = part(sanitizeDocxForWord2007(buf), 'word/numbering.xml')
    // The bidi level justification is corrected...
    expect(numAfter).not.toContain('w:val="start"')
    expect(numAfter).not.toContain('w:val="end"')
    expect(numAfter).toContain('<w:lvlJc w:val="left"')
    // ...but every numbering start value is left exactly as-is.
    expect((numAfter.match(/<w:start w:val="\d+"\s*\/>/g) ?? []).length).toBe(startVals)
  })

  it('keeps document.xml well-formed and still fillable after sanitising', () => {
    const buf = readFileSync(TENDER)
    const labels = findPlaceholdersInDocx(buf)
    const resolved: PlaceholderMatch[] = labels.map((l) => ({ label: l, column: l, score: 1 }))
    const row = Object.fromEntries(labels.map((l) => [l, `val-${l}`]))
    const filled = fillPlaceholdersInDocx(buf, resolved, row)
    const out = sanitizeDocxForWord2007(filled)
    const xml = part(out, 'word/document.xml')
    expect(xml).not.toContain('<w:start ')
    expect(xml).toContain('val-Name of the work')
    // A filled placeholder value is present and the parse round-trips.
    expect(() => new PizZip(out).file('word/document.xml')!.asText()).not.toThrow()
  })

  it('is a no-op (returns a doc with no border start/end) even if run twice', () => {
    const buf = readFileSync(TENDER)
    const once = sanitizeDocxForWord2007(buf)
    const twice = sanitizeDocxForWord2007(once)
    expect(part(twice, 'word/document.xml')).not.toContain('<w:start ')
    expect(part(twice, 'word/document.xml')).not.toContain('<w:end ')
  })
})

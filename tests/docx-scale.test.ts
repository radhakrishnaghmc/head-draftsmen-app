import { describe, it, expect } from 'vitest'
import { readFileSync } from 'fs'
import PizZip from 'pizzip'
import { scaleDocxText } from '../core/docxScale'

const part = (b64: string, name: string): string => new PizZip(b64, { base64: true }).file(name)!.asText()
const sizes = (xml: string): number[] => [...xml.matchAll(/<w:sz w:val="(\d+)"/g)].map((m) => Number(m[1]))

describe('scaleDocxText on the Agreement Bond templates', () => {
  const bond = readFileSync('resources/agreement-template.docx').toString('base64')
  const kompally = readFileSync('resources/agreement-template-kompally.docx').toString('base64')

  it('shrinks body text sizes by the factor', () => {
    const before = sizes(part(bond, 'word/document.xml'))
    const after = sizes(part(scaleDocxText(bond, 0.9), 'word/document.xml'))
    expect(after.length).toBe(before.length)
    const sum = (a: number[]): number => a.reduce((x, y) => x + y, 0)
    expect(sum(after)).toBeLessThan(sum(before))
  })

  it('leaves the blank letterhead-gap lines above the title untouched', () => {
    const body = (b64: string) => part(b64, 'word/document.xml')
    const gap = (xml: string) => {
      const firstText = xml.search(/<w:t(?:\s[^>]*)?>[^<]+<\/w:t>/)
      return xml.slice(xml.indexOf('<w:body>'), xml.lastIndexOf('<w:p ', firstText)).match(/<w:sz w:val="\d+"/g)
    }
    expect(gap(body(scaleDocxText(bond, 0.8)))).toEqual(gap(body(bond)))
  })

  it('never writes a size below 8pt', () => {
    expect(Math.min(...sizes(part(scaleDocxText(bond, 0.1), 'word/document.xml')))).toBeGreaterThanOrEqual(16)
  })

  it("writes Word's implicit 10pt default in when no size is set anywhere (factor 1 changes nothing else)", () => {
    const styles = part(scaleDocxText(kompally, 1), 'word/styles.xml')
    expect(/<w:rPrDefault>[\s\S]*<w:sz w:val="20"\/>[\s\S]*<\/w:rPrDefault>/.test(styles)).toBe(true)
    expect(part(scaleDocxText(kompally, 1), 'word/document.xml')).toBe(part(kompally, 'word/document.xml'))
  })
})

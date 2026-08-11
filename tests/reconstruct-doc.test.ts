import { describe, expect, it } from 'vitest'
import PizZip from 'pizzip'
import { geometryToBlocks, geometryHasText, textLinesToBlocks, type PageGeometry } from '../core/reconstructDoc'
import { buildDocx, type DocTable } from '../core/docxBuilder'

// A tiny 2-column × 2-row bordered table with a centred bold title above it.
const page: PageGeometry = {
  width: 350,
  height: 500,
  texts: [
    { str: 'TITLE', x: 155, y: 60, w: 40, size: 12, bold: true }, // centre ~175 = width/2
    { str: 'Name', x: 55, y: 110, w: 40, size: 11, bold: false },
    { str: 'Ravi', x: 155, y: 110, w: 40, size: 11, bold: false },
    { str: 'Amount', x: 55, y: 140, w: 55, size: 11, bold: false }
    // second-row value cell deliberately empty
  ],
  hlines: [
    { y: 100, x1: 50, x2: 300 },
    { y: 130, x1: 50, x2: 300 },
    { y: 160, x1: 50, x2: 300 }
  ],
  vlines: [
    { x: 50, y1: 100, y2: 160 },
    { x: 150, y1: 100, y2: 160 },
    { x: 300, y1: 100, y2: 160 }
  ]
}

describe('geometryToBlocks', () => {
  const blocks = geometryToBlocks([page])

  it('produces a centred bold title paragraph then a table', () => {
    expect(blocks[0]).toMatchObject({ kind: 'paragraph', align: 'center', runs: [{ text: 'TITLE', bold: true }] })
    const table = blocks.find((b) => b.kind === 'table') as DocTable
    expect(table).toBeDefined()
    expect(table.rows.length).toBe(2)
    expect(table.rows[0].length).toBe(2)
    // Row 1: Name | Ravi
    expect(table.rows[0][0].lines[0][0].text).toBe('Name')
    expect(table.rows[0][1].lines[0][0].text).toBe('Ravi')
    // Row 2 value cell empty
    expect(table.rows[1][1].lines.length).toBe(0)
  })

  it('geometryHasText reflects extractable text', () => {
    expect(geometryHasText([page])).toBe(true)
    expect(geometryHasText([{ width: 1, height: 1, texts: [], hlines: [], vlines: [] }])).toBe(false)
  })
})

describe('buildDocx (Word-valid package)', () => {
  const buf = buildDocx(geometryToBlocks([page]))
  const zip = new PizZip(buf)
  const doc = zip.file('word/document.xml')!.asText()

  it('has no directory entries in the package', () => {
    const dirs = Object.keys(zip.files).filter((n) => (zip.files[n] as { dir?: boolean }).dir)
    expect(dirs).toEqual([])
  })

  it('body ends with sectPr and every table cell has a paragraph', () => {
    expect(/<\/w:sectPr>\s*<\/w:body>/.test(doc)).toBe(true)
    const cells = doc.match(/<w:tc>.*?<\/w:tc>/gs) ?? []
    expect(cells.length).toBeGreaterThan(0)
    expect(cells.filter((c) => !c.includes('<w:p')).length).toBe(0)
  })

  it('builds a valid package from plain OCR lines too', () => {
    const b = buildDocx(textLinesToBlocks(['Line one', '', 'Line two']))
    const d = new PizZip(b).file('word/document.xml')!.asText()
    expect(d).toContain('Line one')
    expect(/<\/w:sectPr>\s*<\/w:body>/.test(d)).toBe(true)
  })

  it('strips XML-forbidden control characters from text (Word rejects them)', () => {
    const b = buildDocx([{ kind: 'paragraph', runs: [{ text: 'Bad\x02char\x1Fend\ttab' }] }])
    const d = new PizZip(b).file('word/document.xml')!.asText()
    // eslint-disable-next-line no-control-regex
    expect(/[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(d)).toBe(false)
    expect(d).toContain('Badcharend\ttab')
  })
})

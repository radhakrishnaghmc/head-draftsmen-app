import { describe, expect, it } from 'vitest'
import { layoutToBlocks, layoutToRows, type OcrPage } from '../core/ocrReconstruct'
import { buildDocx, type DocTable } from '../core/docxBuilder'
import PizZip from 'pizzip'

// A simple 2-column form: label column at x≈40, value column at x≈300, three rows.
const page: OcrPage = {
  lines: [
    { text: 'Name of the work', x: 40, y: 100, w: 180, h: 14 },
    { text: 'Road at Nizampet', x: 300, y: 100, w: 200, h: 14 },
    { text: 'Estimate amount', x: 40, y: 130, w: 160, h: 14 },
    { text: 'Rs 15,00,000/-', x: 300, y: 130, w: 150, h: 14 },
    { text: 'Agency', x: 40, y: 160, w: 90, h: 14 }
    // value for row 3 deliberately missing
  ]
}

describe('layoutToBlocks (image-based reconstruction)', () => {
  const blocks = layoutToBlocks([page])
  const table = blocks.find((b) => b.kind === 'table') as DocTable

  it('rebuilds a 2-column table with rows aligned by y', () => {
    expect(table).toBeDefined()
    expect(table.rows.length).toBe(3)
    expect(table.rows[0][0].lines[0][0].text).toBe('Name of the work')
    expect(table.rows[0][1].lines[0][0].text).toBe('Road at Nizampet')
    expect(table.rows[1][0].lines[0][0].text).toBe('Estimate amount')
    expect(table.rows[1][1].lines[0][0].text).toBe('Rs 15,00,000/-')
    // Row 3: label present, value cell empty
    expect(table.rows[2][0].lines[0][0].text).toBe('Agency')
    expect(table.rows[2][1].lines.length).toBe(0)
  })

  it('the reconstructed table builds a Word-valid docx', () => {
    const doc = new PizZip(buildDocx(blocks)).file('word/document.xml')!.asText()
    expect(doc).toContain('Road at Nizampet')
    expect(/<\/w:sectPr>\s*<\/w:body>/.test(doc)).toBe(true)
  })
})

describe('layoutToRows (image-based Excel grid)', () => {
  it('places cells into columns per visual row', () => {
    const rows = layoutToRows([page])
    expect(rows[0]).toEqual(['Name of the work', 'Road at Nizampet'])
    expect(rows[1]).toEqual(['Estimate amount', 'Rs 15,00,000/-'])
    expect(rows[2][0]).toBe('Agency')
  })
})

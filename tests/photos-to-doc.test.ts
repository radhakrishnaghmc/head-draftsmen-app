import { describe, expect, it } from 'vitest'
import * as XLSX from 'xlsx'
import { splitLines, textToParagraphsHtml, textToRows, buildPhotosWorkbook } from '../core/photosToDoc'

describe('splitLines', () => {
  it('normalises CRLF and trims trailing whitespace, keeping blank lines', () => {
    expect(splitLines('a  \r\nb\r\n\r\nc')).toEqual(['a', 'b', '', 'c'])
  })
})

describe('textToParagraphsHtml', () => {
  it('makes one paragraph per line and escapes HTML, blanks become empty paragraphs', () => {
    expect(textToParagraphsHtml('Hello <b>x</b>\n\nWorld & co')).toBe(
      '<p>Hello &lt;b&gt;x&lt;/b&gt;</p><p>&nbsp;</p><p>World &amp; co</p>'
    )
  })

  it('never emits an empty body for empty input', () => {
    expect(textToParagraphsHtml('')).toBe('<p>&nbsp;</p>')
  })
})

describe('textToRows', () => {
  it('drops blank lines and splits each line into columns on 2+ spaces', () => {
    expect(textToRows('Item   Qty   Rate\n\nCement    10    350')).toEqual([
      ['Item', 'Qty', 'Rate'],
      ['Cement', '10', '350']
    ])
  })

  it('keeps a prose line (single spaces) as one cell', () => {
    expect(textToRows('This is a sentence.')).toEqual([['This is a sentence.']])
  })
})

describe('buildPhotosWorkbook', () => {
  it('produces a readable .xlsx with one row per line and no invented header', () => {
    const buf = buildPhotosWorkbook('Item   Qty\nCement   10', 'OCR')
    const wb = XLSX.read(buf, { type: 'buffer' })
    const sheet = wb.Sheets[wb.SheetNames[0]]
    const aoa = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1 })
    expect(aoa[0]).toEqual(['Item', 'Qty'])
    expect(aoa[1]).toEqual(['Cement', '10'])
  })
})

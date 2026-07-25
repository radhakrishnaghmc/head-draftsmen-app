import { describe, expect, it } from 'vitest'
import { extractLabeledLine } from '../core/ocrLabels'

describe('extractLabeledLine', () => {
  it('reads the value right after the label on the same line', () => {
    const lines = ['Some heading', 'Date: 12.07.2026', 'Other text']
    expect(extractLabeledLine(lines, 'Date')).toBe('12.07.2026')
  })

  it('tolerates OCR spacing/case differences in the label', () => {
    const lines = ['name  OF the  Work : Laying of UGD at Nizampet']
    expect(extractLabeledLine(lines, 'Name of the work')).toBe('Laying of UGD at Nizampet')
  })

  it('falls back to the next non-blank line when the label has nothing after it', () => {
    const lines = ['Wincode', '', 'GHMC-2026-057']
    expect(extractLabeledLine(lines, 'Wincode')).toBe('GHMC-2026-057')
  })

  it('handles a dash separator instead of a colon', () => {
    const lines = ['Circle - Gajularamaram']
    expect(extractLabeledLine(lines, 'Circle')).toBe('Gajularamaram')
  })

  it('strips surrounding quotes from the extracted value', () => {
    const lines = ['Agency: "Radha Krishna Constructions"']
    expect(extractLabeledLine(lines, 'Agency')).toBe('Radha Krishna Constructions')
  })

  it('returns undefined when no line mentions the label at all', () => {
    const lines = ['Random line one', 'Random line two']
    expect(extractLabeledLine(lines, 'Date of Inspection')).toBeUndefined()
  })

  it('does not confuse a label that is a substring of another word', () => {
    const lines = ['Datebook entry: not a date']
    // "Date" alone shouldn't match "Datebook" as if it were the label followed by "book: ..."
    expect(extractLabeledLine(lines, 'Date')).toBeUndefined()
  })
})

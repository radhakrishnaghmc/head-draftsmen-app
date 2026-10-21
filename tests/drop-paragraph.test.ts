import { describe, it, expect } from 'vitest'
import PizZip from 'pizzip'
import { readFileSync } from 'fs'
import { fillPlaceholdersInAllParts, extractAllText, findPlaceholdersInAllParts } from '../core/docx-edit'
import { DROP_PARAGRAPH_MARKER } from '../core/dropMarker'

describe('drop-marked paragraphs', () => {
  it('removes the Diploma line from the Eligibility Criteria when its count resolves to the marker', () => {
    const buf = readFileSync('resources/eligibility-criteria-template.docx')
    const labels = findPlaceholdersInAllParts(buf)
    const resolved = labels.map((label) => ({ label, column: label, score: 1 }))
    const row: Record<string, string> = Object.fromEntries(labels.map((l) => [l, 'X']))
    row['Graduate Engineers'] = '1 No.'
    row['Diploma Engineers'] = DROP_PARAGRAPH_MARKER
    const out = fillPlaceholdersInAllParts(buf, resolved, row)
    const text = extractAllText(out)
    expect(text).toContain('Graduate Engineer (Civil) - 1 No.')
    expect(text).not.toContain('Diploma Engineer')
    expect(text).not.toContain(DROP_PARAGRAPH_MARKER)
    expect(new PizZip(out).file('word/document.xml')).toBeTruthy()
  })
})

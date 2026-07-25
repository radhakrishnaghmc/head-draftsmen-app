import { describe, it, expect } from 'vitest'
import PizZip from 'pizzip'
import {
  listParagraphs,
  setParagraphText,
  applyParagraphEdits,
  findPlaceholdersInDocx,
  fillPlaceholdersInDocx,
  bakeFixedPlaceholdersInDocx
} from '../core/docx-edit'
import type { PlaceholderMatch } from '../core/createDocument'

interface RunSpec {
  text: string
  bold?: boolean
}

/**
 * Builds a minimal in-memory .docx containing only word/document.xml — the
 * only part core/docx-edit.ts's loadDoc() actually reads, so a fixture
 * doesn't need every other OOXML part a real Word file would have.
 */
function buildDocx(paragraphs: RunSpec[][]): Buffer {
  const body = paragraphs
    .map((runs) => {
      const runsXml = runs
        .map((r) => {
          const rPr = r.bold ? '<w:rPr><w:b/></w:rPr>' : ''
          return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(r.text)}</w:t></w:r>`
        })
        .join('')
      return `<w:p>${runsXml}</w:p>`
    })
    .join('')
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`
  const zip = new PizZip()
  zip.file('word/document.xml', xml)
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Whether the run carrying `text` in the document XML is bold (`<w:b/>` in its rPr) — used to verify formatting survived an edit. */
function isRunBold(buffer: Buffer, text: string): boolean {
  const xml = new PizZip(buffer).file('word/document.xml')!.asText()
  const runRe = new RegExp(`<w:r>(?:<w:rPr>([^]*?)</w:rPr>)?<w:t[^>]*>${text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}</w:t></w:r>`)
  const m = runRe.exec(xml)
  return !!m && !!m[1] && m[1].includes('<w:b/>')
}

describe('listParagraphs', () => {
  it('returns the plain text of every paragraph, in order', () => {
    const buffer = buildDocx([[{ text: 'First' }], [{ text: 'Second' }]])
    expect(listParagraphs(buffer)).toEqual(['First', 'Second'])
  })
})

describe('setParagraphText', () => {
  it('preserves an unrelated bold run when only the plain run changes', () => {
    const buffer = buildDocx([[{ text: 'Dear ' }, { text: 'Sir', bold: true }, { text: ', regards' }]])
    const out = setParagraphText(buffer, 0, 'Hello Sir, regards')
    expect(listParagraphs(out)).toEqual(['Hello Sir, regards'])
    expect(isRunBold(out, 'Sir')).toBe(true)
  })

  it('returns the same content when newText equals the original', () => {
    const buffer = buildDocx([[{ text: 'Unchanged' }]])
    const out = setParagraphText(buffer, 0, 'Unchanged')
    expect(listParagraphs(out)).toEqual(['Unchanged'])
  })
})

describe('applyParagraphEdits', () => {
  it('rewrites several paragraphs in one pass', () => {
    const buffer = buildDocx([[{ text: 'One' }], [{ text: 'Two' }], [{ text: 'Three' }]])
    const out = applyParagraphEdits(buffer, [
      { index: 0, newText: 'ONE' },
      { index: 2, newText: 'THREE' }
    ])
    expect(listParagraphs(out)).toEqual(['ONE', 'Two', 'THREE'])
  })

  it('is a no-op when no edit changes anything', () => {
    const buffer = buildDocx([[{ text: 'Same' }]])
    const out = applyParagraphEdits(buffer, [{ index: 0, newText: 'Same' }])
    expect(listParagraphs(out)).toEqual(['Same'])
  })
})

describe('findPlaceholdersInDocx', () => {
  it('finds distinct placeholders across paragraphs, in first-seen order', () => {
    const buffer = buildDocx([[{ text: 'Name: {{Name of the work}}' }], [{ text: 'Dated {{Agmt Date}}' }]])
    expect(findPlaceholdersInDocx(buffer)).toEqual(['Name of the work', 'Agmt Date'])
  })

  it('dedupes a placeholder that appears more than once', () => {
    const buffer = buildDocx([[{ text: '{{Circle}} ... {{Circle}}' }]])
    expect(findPlaceholdersInDocx(buffer)).toEqual(['Circle'])
  })

  it('reads a placeholder split across several runs as one clean label', () => {
    // Mirrors the real-world case in core/createDocument.ts's tests: Word
    // frequently splits one conceptual run across several <w:r> elements
    // (spell-check, language tagging) even mid-placeholder.
    const buffer = buildDocx([[{ text: '{{Amount ' }, { text: 'of ', bold: true }, { text: 'Estimate}}' }]])
    expect(findPlaceholdersInDocx(buffer)).toEqual(['Amount of Estimate'])
  })

  it('returns an empty array when there are no placeholders', () => {
    expect(findPlaceholdersInDocx(buildDocx([[{ text: 'Just plain text.' }]]))).toEqual([])
  })
})

describe('fillPlaceholdersInDocx', () => {
  it('replaces a resolved placeholder with the row value', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Name of the work', column: 'Name of Work', score: 0.9 }]
    const row = { 'Name of Work': 'Road repair, Ward 12' }
    const buffer = buildDocx([[{ text: '{{Name of the work}}' }]])
    expect(listParagraphs(fillPlaceholdersInDocx(buffer, resolved, row))).toEqual(['Road repair, Ward 12'])
  })

  it('blanks an unresolved placeholder rather than leaving the token in place', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Mystery Field', column: null, score: 0 }]
    const buffer = buildDocx([[{ text: 'Value: {{Mystery Field}}.' }]])
    expect(listParagraphs(fillPlaceholdersInDocx(buffer, resolved, {}))).toEqual(['Value: .'])
  })

  it('replaces every occurrence of a repeated placeholder', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Name of the work', column: 'Name of Work', score: 0.9 }]
    const row = { 'Name of Work': 'Bridge work' }
    const buffer = buildDocx([[{ text: '{{Name of the work}} ... {{Name of the work}}' }]])
    expect(listParagraphs(fillPlaceholdersInDocx(buffer, resolved, row))).toEqual(['Bridge work ... Bridge work'])
  })

  it('leaves an unrelated bold run in the same paragraph untouched', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Name', column: 'Name', score: 0.9 }]
    const buffer = buildDocx([[{ text: '{{Name}} — ' }, { text: 'IMPORTANT', bold: true }]])
    const out = fillPlaceholdersInDocx(buffer, resolved, { Name: 'Road repair' })
    expect(listParagraphs(out)).toEqual(['Road repair — IMPORTANT'])
    expect(isRunBold(out, 'IMPORTANT')).toBe(true)
  })

  it('leaves paragraphs with no placeholder untouched', () => {
    const resolved: PlaceholderMatch[] = [{ label: 'Name', column: 'Name', score: 0.9 }]
    const buffer = buildDocx([[{ text: '{{Name}}' }], [{ text: 'Plain paragraph, no braces.' }]])
    const out = fillPlaceholdersInDocx(buffer, resolved, { Name: 'X' })
    expect(listParagraphs(out)).toEqual(['X', 'Plain paragraph, no braces.'])
  })
})

describe('bakeFixedPlaceholdersInDocx', () => {
  it('replaces only the given labels, matched case-insensitively', () => {
    const buffer = buildDocx([[{ text: '{{Circle}} / {{circle}} / {{Zone}} / {{CNO}}' }]])
    const out = bakeFixedPlaceholdersInDocx(buffer, {
      zone: 'Cyberabad',
      circle: 'Gajularamaram Circle-57',
      cno: '57'
    })
    expect(listParagraphs(out)).toEqual(['Gajularamaram Circle-57 / Gajularamaram Circle-57 / Cyberabad / 57'])
  })

  it('leaves every other placeholder untouched for later per-row resolution', () => {
    const buffer = buildDocx([[{ text: '{{Circle}} — {{Name of the work}} — {{Estimate Amount}}' }]])
    const out = bakeFixedPlaceholdersInDocx(buffer, { circle: 'Gajularamaram Circle-57' })
    expect(listParagraphs(out)).toEqual(['Gajularamaram Circle-57 — {{Name of the work}} — {{Estimate Amount}}'])
  })

  it('is idempotent — running it again after the label is gone changes nothing', () => {
    const buffer = buildDocx([[{ text: '{{Zone}}' }]])
    const once = bakeFixedPlaceholdersInDocx(buffer, { zone: 'Cyberabad' })
    const twice = bakeFixedPlaceholdersInDocx(once, { zone: 'Cyberabad' })
    expect(listParagraphs(twice)).toEqual(['Cyberabad'])
  })
})

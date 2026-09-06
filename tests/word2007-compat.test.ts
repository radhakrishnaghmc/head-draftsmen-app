import { describe, it, expect } from 'vitest'
import { readFileSync } from 'fs'
import { resolve } from 'path'
import PizZip from 'pizzip'
import { fillPlaceholdersInDocx, findPlaceholdersInDocx } from '../core/docx-edit'
import { sanitizeDocxForWord2007 } from '../core/word2007Compat'
import type { PlaceholderMatch } from '../core/createDocument'
import type { Element } from '@xmldom/xmldom'

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

  it('reorders out-of-sequence <w:pPr> children on a template-authored (not html-to-docx) .docx', () => {
    // Found via a live "SE docs won't open in Word 2007" report: several
    // bundled LibreOffice templates (this one included — 16 real violations)
    // have <w:pPr> children out of the OOXML CT_PPr sequence, independently
    // of the bidirectional-construct bug above. This was previously only
    // fixed for html-to-docx output (sanitizeHtmlDocxForWord2007) — templates
    // filled via docxtemplater got no equivalent treatment until now.
    const buf = readFileSync(resolve(__dirname, '../resources/eligibility-criteria-template.docx'))
    const PPR_ORDER = [
      'w:pStyle','w:keepNext','w:keepLines','w:pageBreakBefore','w:framePr','w:widowControl','w:numPr',
      'w:suppressLineNumbers','w:pBdr','w:shd','w:tabs','w:suppressAutoHyphens','w:kinsoku','w:wordWrap',
      'w:overflowPunct','w:topLinePunct','w:autoSpaceDE','w:autoSpaceDN','w:bidi','w:adjustRightInd',
      'w:snapToGrid','w:spacing','w:ind','w:contextualSpacing','w:mirrorIndents','w:suppressOverlap','w:jc',
      'w:textDirection','w:textAlignment','w:textboxTightWrap','w:outlineLvl','w:divId','w:cnfStyle','w:rPr',
      'w:sectPr','w:pPrChange'
    ]
    const countOutOfOrder = (xmlText: string): number => {
      const xml = new DOMParser().parseFromString(xmlText, 'text/xml')
      const pPrs = Array.from(xml.getElementsByTagName('w:pPr')) as unknown as Element[]
      let bad = 0
      for (const pPr of pPrs) {
        const direct = (Array.from(pPr.childNodes) as unknown as Element[]).filter((n) => n.nodeType === 1)
        const ranks = direct.map((c) => {
          const i = PPR_ORDER.indexOf(c.nodeName)
          return i === -1 ? 999 : i
        })
        for (let i = 1; i < ranks.length; i++) {
          if (ranks[i] < ranks[i - 1]) {
            bad++
            break
          }
        }
      }
      return bad
    }

    expect(countOutOfOrder(part(buf, 'word/document.xml'))).toBeGreaterThan(0)
    const out = sanitizeDocxForWord2007(buf)
    expect(countOutOfOrder(part(out, 'word/document.xml'))).toBe(0)
  })

  it('strips explicit directory entries from the zip — the real cause of "unreadable content" that survived the bidi and child-order fixes', () => {
    // Real bug, found live: a downloaded Agreement Bond ("Agreement Bond -
    // Ganta Narasimha Rao.docx") still showed Word's "unreadable content"
    // prompt in Office 365 too, not just Word 2007 — proving the bidi/
    // child-order fixes above (both Word-2007-specific schema strictness)
    // were never going to fix this one, because it's a different corruption
    // class entirely: resources/agreement-template.docx has 5 explicit
    // directory entries (_rels/, docProps/, word/, word/theme/, word/_rels/)
    // baked into the raw template, dated separately from its other parts —
    // added by some later re-export/re-zip tool. Word rejects a package with
    // these in EVERY version (not just 2007), which is exactly why the
    // Word-2007-only fixes didn't help.
    const buf = readFileSync(resolve(__dirname, '../resources/agreement-template.docx'))
    const zipBefore = new PizZip(buf)
    const dirsBefore = Object.keys(zipBefore.files).filter((f) => (zipBefore.files as any)[f].dir)
    expect(dirsBefore.length).toBeGreaterThan(0)

    const out = sanitizeDocxForWord2007(buf)
    const zipAfter = new PizZip(out)
    const dirsAfter = Object.keys(zipAfter.files).filter((f) => (zipAfter.files as any)[f].dir)
    expect(dirsAfter).toEqual([])
    // Every real file must survive the folder-entry removal (only the
    // zero-length directory marker entries should be dropped).
    expect(zipAfter.file('word/document.xml')).not.toBeNull()
    expect(zipAfter.file('word/styles.xml')).not.toBeNull()
    expect(zipAfter.file('word/_rels/document.xml.rels')).not.toBeNull()
  })

  it('fixes the invalid <w:sz-cs> element name — the ACTUAL cause of the SE Agreement Bond still failing after the directory-entry fix', () => {
    // Real bug, continued: directory-entries turned out to be a DIFFERENT
    // template (agreement-template.docx, the Tools tab's EE "Agreement
    // Bond") than the one the user was actually generating — the SE zonal
    // workflow's (since-removed) se-agreement-bond-template.docx. Re-
    // reproduced live through the real SE Agreement/Work Order workflow
    // (uploaded the same L-1 + Intimation, downloaded the real Agreement
    // Bond tile) and found <w:sz-cs w:val="20"/> in the output — NOT a valid
    // OOXML element at all (the real one is <w:szCs>, camelCase, no hyphen).
    // This is WELL-FORMED XML (a hyphen in an element name is syntactically
    // legal), so it passed every earlier check (bidi, child-order,
    // directory-entries, XML well-formedness) — only Word's own OOXML
    // *schema* validation catches it, which is why it broke every Word
    // version including Office 365, and why three earlier fixes in a row
    // didn't touch it. Found systemically in 5 SE templates (318 occurrences
    // total): eligibility-criteria (25), se-agreement-bond (41),
    // se-agreement-note (57), se-contract-deed (165), ts-note (30) — all
    // sharing some earlier edit that introduced the typo. The three SE Work
    // Order/Agreement templates were later removed for a from-scratch
    // rebuild, so this regression test now reproduces the bug against
    // eligibility-criteria-template.docx (still bundled, still affected)
    // instead.
    const buf = readFileSync(resolve(__dirname, '../resources/eligibility-criteria-template.docx'))
    expect(part(buf, 'word/document.xml')).toMatch(/<w:sz-cs\b/)

    const out = sanitizeDocxForWord2007(buf)
    const fixedDoc = part(out, 'word/document.xml')
    expect(fixedDoc).not.toMatch(/<w:sz-cs\b/)
    expect(fixedDoc).not.toMatch(/<\/w:sz-cs>/)
    expect(fixedDoc).toContain('<w:szCs w:val="20"/>')
  })

  it('strips embedded fonts from the logo-header template variants — Word 2007 predates font embedding entirely', () => {
    // Real bug, found live: downloaded documents from the new logo-header
    // template variants (Nizampet Circle-58's own letterhead) errored out in
    // older Word. Unlike every fix above, this isn't a bidi/schema-order/
    // typo issue — real Word 2013+ saved these templates with "Embed fonts
    // in the file" on, to carry "Noto Sans Telugu" for the Telugu caption
    // under the state emblem. That bakes a <w:embedRegular> ref into
    // word/fontTable.xml, a relationship in its .rels, and a word/fonts/*
    // .fntdata part — all of which are OOXML's font-embedding extension,
    // which Word 2007's schema has never heard of, so it refuses to open
    // the package at all. Word substitutes a fallback font for a declared
    // name it can't find either way, so dropping the embedding is lossless
    // in practice for this app's use.
    for (const name of [
      'work-order-template-header2.docx',
      'civil-tender-template-header2.docx',
      'file-backer-template-header2.docx',
      'intimation-template-2.docx'
    ]) {
      const buf = readFileSync(resolve(__dirname, '../resources', name))
      const zipBefore = new PizZip(buf)
      const fontParts = Object.keys(zipBefore.files).filter((f) => /^word\/fonts\//.test(f))
      expect(fontParts.length, `${name} should carry an embedded font`).toBeGreaterThan(0)
      expect(part(buf, 'word/fontTable.xml')).toMatch(/<w:embedRegular\b/)

      const out = sanitizeDocxForWord2007(buf)
      const zipAfter = new PizZip(out)
      for (const f of fontParts) expect(zipAfter.file(f), `${f} should be removed`).toBeNull()
      expect(part(out, 'word/fontTable.xml')).not.toMatch(/<w:embed(Regular|Bold|Italic)\b/)
      expect(part(out, 'word/settings.xml')).not.toContain('embedTrueTypeFonts')
    }
  })
})

import { convertHtmlToDocx } from '../core/htmlToDocx'
import { buildNoteSubmittedHtml, noteSubmittedFromRow } from '../core/noteSubmitted'
import { DOMParser } from '@xmldom/xmldom'

function childSeqs(xml: string, container: string): string[] {
  const doc = new DOMParser().parseFromString(xml, 'text/xml')
  const out = new Set<string>()
  for (const n of Array.from(doc.getElementsByTagName(container)) as any[]) {
    const kids = (Array.from(n.childNodes) as any[])
      .filter((c) => c.nodeType === 1)
      .map((c) => c.nodeName)
    if (kids.length) out.add(kids.join(','))
  }
  return [...out]
}

describe('convertHtmlToDocx → Word 2007 (html-to-docx sanitising)', () => {
  it('produces a Word-2007-openable Note Submitted docx', async () => {
    const data = noteSubmittedFromRow(
      {
        'Name of the work': 'Laying of UGD (Reserved for ST)',
        Circle: 'Nizampet',
        'Amount of estimate': '18.00',
        'Tender Notice No': '11/DB/EE',
        'Tender notice Date': '14.07.2026',
        'Name of the Agency': 'L SURENDER',
        Reservation: 'ST'
      },
      {},
      {},
      'Nizampet'
    )
    const buf = await convertHtmlToDocx(buildNoteSubmittedHtml(data))

    // 1. No invalid .rels content-type overrides (Word 2007 rejects them).
    expect(part(buf, '[Content_Types].xml')).not.toMatch(/PartName="[^"]*\.rels"/i)

    // 2. Property containers are in OOXML schema order.
    const doc = part(buf, 'word/document.xml')
    expect(childSeqs(doc, 'w:tblPr')).toEqual(['w:tblW,w:jc,w:tblCellSpacing,w:tblBorders,w:tblCellMar'])
    expect(childSeqs(doc, 'w:tblCellMar')).toEqual(['w:top,w:left,w:bottom,w:right'])
    // pPr may appear as just <w:spacing> and as <w:spacing,w:jc> — never jc-before-spacing.
    expect(childSeqs(doc, 'w:pPr')).not.toContain('w:jc,w:spacing')
  })
})

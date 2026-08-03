import PizZip from 'pizzip'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import type { Document, Element } from '@xmldom/xmldom'

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

/**
 * A LibreOffice-authored .docx writes several things in the *bidirectional*
 * ("start"/"end") form that OOXML gained long after Word 2007. Word 2010+
 * tolerates them; Word 2007's schema does not, and refuses to open the file
 * ("… cannot be opened because there are problems with the contents",
 * Unspecified error pointing at the first offending node). All of them are the
 * logical spelling of a physical left/right and convert losslessly for a
 * left-to-right document:
 *
 *   1. border / cell-margin SIDE ELEMENTS  <w:start>/<w:end>  → <w:left>/<w:right>
 *   2. justification VALUE   <w:jc w:val="start|end">          → "left"/"right"
 *   3. indent ATTRIBUTES     <w:ind w:start=/w:end=/…Chars=>   → w:left/w:right/…
 *
 * This is what broke the Civil Tender Document in Word 2007 (its template
 * carried 152 border <w:start>/<w:end> plus 21 <w:jc w:val="start|end">).
 */

// Border / margin groups whose left/right sides appear as <w:start>/<w:end>.
// A list level's own <w:start w:val="1"/> (a numbering start value) is a direct
// child of <w:lvl>, never of one of these, so scoping the element rename to
// these containers leaves numbering untouched.
const SIDE_CONTAINERS = ['w:tcBorders', 'w:tblBorders', 'w:pBdr', 'w:tblCellMar', 'w:tcMar']

// <w:ind> bidi attributes → their physical equivalents.
const IND_ATTRS: [from: string, to: string][] = [
  ['w:start', 'w:left'],
  ['w:end', 'w:right'],
  ['w:startChars', 'w:leftChars'],
  ['w:endChars', 'w:rightChars']
]

// Parts that can carry any of the above. numbering.xml / styles.xml are
// included because a list level or a paragraph style can also hold a
// <w:jc w:val="start">; the element rename staying container-scoped keeps a
// numbering start value safe there.
const FIXABLE_PART_RE = /^word\/(document|header\d+|footer\d+|footnotes|endnotes|styles|numbering)\.xml$/

// Cheap pre-filter so a part with none of the constructs is never reparsed.
const MAYBE_RE = /<w:start[ />]|<w:end[ />]|w:val="start"|w:val="end"|w:start(Chars)?="|w:end(Chars)?="/

function els(root: Document, name: string): Element[] {
  return Array.from(root.getElementsByTagName(name)) as unknown as Element[]
}

/** Rename <w:start>→<w:left> / <w:end>→<w:right> inside border/margin groups, in place (preserving OOXML child order). */
function fixSideElements(xml: Document): boolean {
  let changed = false
  for (const container of SIDE_CONTAINERS) {
    for (const parent of els(xml, container)) {
      for (const child of Array.from(parent.childNodes) as unknown as Element[]) {
        const to = child.nodeName === 'w:start' ? 'w:left' : child.nodeName === 'w:end' ? 'w:right' : ''
        if (!to) continue
        // Never emit a duplicate physical side if the file already had one.
        if ((Array.from(parent.childNodes) as unknown as Element[]).some((n) => n.nodeName === to)) {
          parent.removeChild(child)
          changed = true
          continue
        }
        const repl = xml.createElementNS(W_NS, to) as unknown as Element
        for (const attr of Array.from(child.attributes)) repl.setAttribute(attr.name, attr.value)
        parent.replaceChild(repl, child)
        changed = true
      }
    }
  }
  return changed
}

// Justification elements that share the ST_Jc enum: paragraph/table <w:jc> and
// a list level's <w:lvlJc> (numbering.xml). Both take the bidi start/end.
const JC_ELEMENTS = ['w:jc', 'w:lvlJc']

/** <w:jc>/<w:lvlJc> w:val="start|end" → "left"/"right". */
function fixJustification(xml: Document): boolean {
  let changed = false
  for (const name of JC_ELEMENTS) {
    for (const jc of els(xml, name)) {
      const v = jc.getAttribute('w:val')
      const to = v === 'start' ? 'left' : v === 'end' ? 'right' : ''
      if (!to) continue
      jc.setAttribute('w:val', to)
      changed = true
    }
  }
  return changed
}

/** <w:ind w:start=/w:end=/w:startChars=/w:endChars=> → the physical attribute names. */
function fixIndentAttrs(xml: Document): boolean {
  let changed = false
  for (const ind of els(xml, 'w:ind')) {
    for (const [from, to] of IND_ATTRS) {
      if (!ind.hasAttribute(from)) continue
      const val = ind.getAttribute(from) ?? ''
      ind.removeAttribute(from)
      if (!ind.hasAttribute(to)) ind.setAttribute(to, val)
      changed = true
    }
  }
  return changed
}

/**
 * Rewrite every Word-2007-incompatible bidirectional construct in a filled
 * .docx so it opens in Word 2007. A no-op — the same buffer is returned — when
 * there's nothing to fix, so it's safe to run over every exported document.
 */
export function sanitizeDocxForWord2007(buffer: Buffer): Buffer {
  const zip = new PizZip(buffer)
  let changed = false
  for (const name of Object.keys(zip.files)) {
    if (!FIXABLE_PART_RE.test(name)) continue
    const xmlText = zip.file(name)?.asText()
    if (!xmlText || !MAYBE_RE.test(xmlText)) continue
    const xml = new DOMParser().parseFromString(xmlText, 'text/xml')
    const partChanged = [fixSideElements(xml), fixJustification(xml), fixIndentAttrs(xml)].some(Boolean)
    if (partChanged) {
      zip.file(name, new XMLSerializer().serializeToString(xml))
      changed = true
    }
  }
  return changed ? zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) : buffer
}

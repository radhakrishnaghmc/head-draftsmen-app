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

// `<w:sz-cs>` is not a valid OOXML element at all (the real one is
// `<w:szCs>`, camelCase, no hyphen) — found baked into 5 SE templates
// (318 occurrences total: eligibility-criteria, ts-note, and three SE
// Work Order/Agreement templates since removed and slated for a from-
// scratch rebuild), all sharing some earlier edit that introduced the
// same typo. This is well-formed XML (hyphens are
// legal in element names), so it passes an XML-syntax check — Word's own
// OOXML *schema* validation is what rejects it, in every Word version,
// which is why this survived the Word-2007-specific bidi/child-order fixes
// AND the directory-entries fix (a different template) unnoticed. A plain
// tag-name string fix, not a DOM operation.
function fixInvalidSzCsTag(xmlText: string): string {
  return xmlText.replace(/<w:sz-cs\b/g, '<w:szCs').replace(/<\/w:sz-cs>/g, '</w:szCs>')
}

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

// Canonical child-element order for each ordered property container (the OOXML
// CT_* sequence). A child not listed keeps its relative position after the known
// ones. Both the physical (left/right) and logical (start/end) side names are
// listed so this is correct whether or not the bidi rename above has run.
const CHILD_ORDER: Record<string, string[]> = {
  'w:pPr': ['w:pStyle','w:keepNext','w:keepLines','w:pageBreakBefore','w:framePr','w:widowControl','w:numPr','w:suppressLineNumbers','w:pBdr','w:shd','w:tabs','w:suppressAutoHyphens','w:kinsoku','w:wordWrap','w:overflowPunct','w:topLinePunct','w:autoSpaceDE','w:autoSpaceDN','w:bidi','w:adjustRightInd','w:snapToGrid','w:spacing','w:ind','w:contextualSpacing','w:mirrorIndents','w:suppressOverlap','w:jc','w:textDirection','w:textAlignment','w:textboxTightWrap','w:outlineLvl','w:divId','w:cnfStyle','w:rPr','w:sectPr','w:pPrChange'],
  'w:rPr': ['w:rStyle','w:rFonts','w:b','w:bCs','w:i','w:iCs','w:caps','w:smallCaps','w:strike','w:dstrike','w:outline','w:shadow','w:emboss','w:imprint','w:noProof','w:snapToGrid','w:vanish','w:webHidden','w:color','w:spacing','w:w','w:kern','w:position','w:sz','w:szCs','w:highlight','w:u','w:effect','w:bdr','w:shd','w:fitText','w:vertAlign','w:rtl','w:cs','w:em','w:lang','w:eastAsianLayout','w:specVanish','w:oMath','w:rPrChange'],
  'w:tblPr': ['w:tblStyle','w:tblpPr','w:tblOverlap','w:bidiVisual','w:tblStyleRowBandSize','w:tblStyleColBandSize','w:tblW','w:jc','w:tblCellSpacing','w:tblInd','w:tblBorders','w:shd','w:tblLayout','w:tblCellMar','w:tblLook','w:tblCaption','w:tblDescription','w:tblPrChange'],
  'w:trPr': ['w:cnfStyle','w:divId','w:gridBefore','w:gridAfter','w:wBefore','w:wAfter','w:cantSplit','w:trHeight','w:tblHeader','w:tblCellSpacing','w:jc','w:hidden','w:ins','w:del','w:trPrChange'],
  'w:tcPr': ['w:cnfStyle','w:tcW','w:gridSpan','w:hMerge','w:vMerge','w:tcBorders','w:shd','w:noWrap','w:tcMar','w:textDirection','w:tcFitText','w:vAlign','w:hideMark','w:headers','w:cellIns','w:cellDel','w:cellMerge','w:tcPrChange'],
  'w:tblCellMar': ['w:top','w:start','w:left','w:bottom','w:end','w:right'],
  'w:tcMar': ['w:top','w:start','w:left','w:bottom','w:end','w:right'],
  'w:tblBorders': ['w:top','w:start','w:left','w:bottom','w:end','w:right','w:insideH','w:insideV'],
  'w:tcBorders': ['w:top','w:start','w:left','w:bottom','w:end','w:right','w:insideH','w:insideV','w:tl2br','w:tr2bl'],
  'w:pBdr': ['w:top','w:left','w:bottom','w:right','w:between','w:bar']
}

/** Stable-reorder a container's element children into `order`; drops the
 * whitespace-only text nodes between them (cosmetic). Returns true if it moved
 * anything. Existing nodes are moved (not recreated), so namespaces survive. */
function reorderChildren(parent: Element, order: string[]): boolean {
  const elems = (Array.from(parent.childNodes) as unknown as Element[]).filter((n) => n.nodeType === 1)
  if (elems.length < 2) return false
  const rank = (n: Element) => {
    const i = order.indexOf(n.nodeName)
    return i === -1 ? order.length : i
  }
  const sorted = elems.map((el, i) => ({ el, i })).sort((a, b) => rank(a.el) - rank(b.el) || a.i - b.i).map((x) => x.el)
  if (sorted.every((el, i) => el === elems[i])) return false
  while (parent.firstChild) parent.removeChild(parent.firstChild)
  for (const el of sorted) parent.appendChild(el)
  return true
}

/**
 * Rewrite every Word-incompatible construct in a filled .docx so it opens
 * cleanly. A no-op — the same buffer is returned — when there's nothing to
 * fix, so it's safe to run over every exported document. Fixes:
 *   1. Bidirectional border/justification/indent constructs (see above) —
 *      Word 2007-specific schema strictness (Word 2010+ tolerates them).
 *   2. Property-container children out of the OOXML CT_* schema sequence —
 *      LibreOffice-authored .docx templates routinely emit these out of
 *      order. Also Word-2007-specific strictness.
 *   3. Explicit directory entries in the zip (e.g. "word/", "_rels/" as
 *      their own zero-length entries, not just implied by file paths) —
 *      found in `resources/agreement-template.docx` (5 of them, dated
 *      separately from the file's other parts — added by some later
 *      re-export/re-zip tool) and `public-participation-book-template.docx`
 *      (6). Unlike (1)/(2), Word REJECTS these in every version, including
 *      Office 365 — a genuine "unreadable content" failure was traced to this
 *      specifically (bidi + child-order fixes alone did not resolve it,
 *      because this is a fundamentally different corruption class: a zip/OPC
 *      package-structure problem, not a schema-strictness one).
 *   4. The invalid `<w:sz-cs>` element name (see fixInvalidSzCsTag) — found
 *      in 5 SE templates, 318 occurrences total. Also rejected in every
 *      Word version (it's simply not a real OOXML element), and — unlike
 *      (1)-(3) — passes an XML well-formedness check, since a hyphen in an
 *      element name is syntactically legal; only OOXML's own schema (which
 *      Word validates against, not a generic XML parser) rejects it.
 * Every FIXABLE_PART_RE part is parsed unconditionally (not gated behind
 * MAYBE_RE) because a child-order problem can exist with none of the bidi
 * patterns present.
 */
export function sanitizeDocxForWord2007(buffer: Buffer): Buffer {
  const zip = new PizZip(buffer)
  let changed = false
  for (const name of Object.keys(zip.files)) {
    if (!FIXABLE_PART_RE.test(name)) continue
    let xmlText = zip.file(name)?.asText()
    if (!xmlText) continue
    const fixedTag = fixInvalidSzCsTag(xmlText)
    let partChanged = fixedTag !== xmlText
    xmlText = fixedTag
    const xml = new DOMParser().parseFromString(xmlText, 'text/xml')
    if ([fixSideElements(xml), fixJustification(xml), fixIndentAttrs(xml)].some(Boolean)) partChanged = true
    for (const [container, order] of Object.entries(CHILD_ORDER)) {
      for (const node of els(xml, container)) if (reorderChildren(node, order)) partChanged = true
    }
    if (partChanged) {
      zip.file(name, new XMLSerializer().serializeToString(xml))
      changed = true
    }
  }

  // Delete only the folder ENTRY from the map (zip.remove() would cascade
  // and delete the files inside the folder too — e.g. removing "word/"
  // would drop word/document.xml, which is a separate key and must survive).
  const files = zip.files as Record<string, { dir?: boolean }>
  for (const name of Object.keys(files)) {
    if (files[name]?.dir) {
      delete files[name]
      changed = true
    }
  }

  return changed ? zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) : buffer
}

// ── html-to-docx → Word 2007 ──────────────────────────────────────────────
// The html-to-docx library emits OOXML that Word 2010+ opens but Word 2007
// rejects ("… problems with the contents"), for two reasons this fixes:
//   1. [Content_Types].xml carries invalid <Override> entries for the .rels
//      relationship parts — those are typed only via the `rels` Default, and
//      Word 2007's stricter OPC reader refuses the package over the duplicate.
//   2. Property containers (<w:pPr>, <w:tblPr>, <w:tblCellMar>, …) come out with
//      their child elements in the WRONG order — e.g. tblPr as
//      "tblBorders,tblCellSpacing,tblW,tblCellMar,jc" — and Word 2007 validates
//      the CT_* child SEQUENCE strictly (Word 2010+ is lenient). Reordering to
//      the schema sequence is lossless. (CHILD_ORDER/reorderChildren for this
//      now live above, shared with sanitizeDocxForWord2007 — template-authored
//      LibreOffice .docx files have the exact same child-order problem.)

/**
 * Move the body-level <w:sectPr> to be the LAST child of <w:body>. html-to-docx
 * emits it as the FIRST child, but the OOXML CT_Body sequence requires the
 * section properties to come last — Word refuses to open the file otherwise
 * ("Word experienced an error trying to open the file"), even though LibreOffice
 * and python-docx tolerate it. Only the direct body child is moved (a sectPr
 * nested in a paragraph's pPr is a mid-document section break, left alone).
 */
function fixBodySectPr(xml: Document): boolean {
  const body = els(xml, 'w:body')[0]
  if (!body) return false
  let sect: Element | undefined
  for (let n = body.firstChild; n; n = n.nextSibling) {
    if ((n as Element).nodeName === 'w:sectPr') {
      sect = n as Element
      break
    }
  }
  if (!sect || body.lastChild === sect) return false
  body.removeChild(sect as never)
  body.appendChild(sect as never)
  return true
}

/** Remove the invalid <Override> entries for .rels parts from [Content_Types].xml. */
function fixContentTypesRels(xml: Document): boolean {
  let changed = false
  for (const ov of els(xml, 'Override')) {
    if (/\.rels$/i.test(ov.getAttribute('PartName') ?? '')) {
      ov.parentNode?.removeChild(ov as never)
      changed = true
    }
  }
  return changed
}

/**
 * Make an html-to-docx buffer open in Word 2007: strip the invalid .rels
 * content-type overrides and reorder every property container's children into
 * the OOXML schema sequence. Also runs the bidirectional fixes above (a no-op
 * for html-to-docx output, kept for safety). Returns the same buffer unchanged
 * when there's nothing to fix.
 */
export function sanitizeHtmlDocxForWord2007(buffer: Buffer): Buffer {
  const zip = new PizZip(buffer)
  let changed = false

  const ctName = '[Content_Types].xml'
  const ctText = zip.file(ctName)?.asText()
  if (ctText && /PartName="[^"]*\.rels"/i.test(ctText)) {
    const ctXml = new DOMParser().parseFromString(ctText, 'text/xml')
    if (fixContentTypesRels(ctXml)) {
      zip.file(ctName, new XMLSerializer().serializeToString(ctXml))
      changed = true
    }
  }

  for (const name of Object.keys(zip.files)) {
    if (!FIXABLE_PART_RE.test(name)) continue
    const xmlText = zip.file(name)?.asText()
    if (!xmlText) continue
    const xml = new DOMParser().parseFromString(xmlText, 'text/xml')
    // Note: each fix must run (not short-circuit), so call them separately.
    let partChanged = fixSideElements(xml)
    if (fixJustification(xml)) partChanged = true
    if (fixIndentAttrs(xml)) partChanged = true
    if (fixBodySectPr(xml)) partChanged = true
    for (const [container, order] of Object.entries(CHILD_ORDER)) {
      for (const node of els(xml, container)) if (reorderChildren(node, order)) partChanged = true
    }
    if (partChanged) {
      zip.file(name, new XMLSerializer().serializeToString(xml))
      changed = true
    }
  }

  // html-to-docx's zip includes explicit directory entries (_rels/, word/,
  // word/theme/, …). Microsoft Word — notably on macOS — refuses to open a
  // .docx whose package contains folder entries ("Word experienced an error
  // trying to open the file"), even though LibreOffice and python-docx accept
  // it. Real Word/PizZip packages have none, so drop them and regenerate.
  let removedDir = false
  const files = zip.files as Record<string, { dir?: boolean }>
  for (const name of Object.keys(files)) {
    // Delete only the folder ENTRY from the map — zip.remove() would cascade and
    // delete the files inside the folder too (e.g. removing "word/" drops
    // word/document.xml). The contained files are separate keys, left intact.
    if (files[name]?.dir) {
      delete files[name]
      removedDir = true
    }
  }
  if (removedDir) changed = true

  return changed ? zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) : buffer
}

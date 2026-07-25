import PizZip from 'pizzip'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import type { Document, Element } from '@xmldom/xmldom'
import { PLACEHOLDER_RE } from './createDocument'
import type { PlaceholderMatch } from './createDocument'

const DOC_XML = 'word/document.xml'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function loadDoc(buffer: Buffer, part: string = DOC_XML): { zip: PizZip; xml: Document } {
  const zip = new PizZip(buffer)
  const xmlText = zip.file(part)?.asText()
  if (!xmlText) throw new Error(`${part} not found in .docx`)
  const xml = new DOMParser().parseFromString(xmlText, 'text/xml')
  return { zip, xml }
}

function tag(el: Element | Document, name: string): Element[] {
  return Array.from(el.getElementsByTagName(name)) as unknown as Element[]
}

/** Plain text of a run element (concatenation of its <w:t> children). */
function runText(run: Element): string {
  return tag(run, 'w:t')
    .map((t) => t.textContent ?? '')
    .join('')
}

/**
 * Plain text of every body paragraph, in document order (includes paragraphs
 * inside table cells). Mirrors the order docx-preview renders `article p`,
 * so a rendered paragraph index maps to the same entry here.
 */
export function listParagraphs(buffer: Buffer, part: string = DOC_XML): string[] {
  const { xml } = loadDoc(buffer, part)
  return tag(xml, 'w:p').map((p) =>
    tag(p, 'w:t')
      .map((t) => t.textContent ?? '')
      .join('')
  )
}

/**
 * Resolve the true paragraph index to operate on. The renderer supplies a
 * numeric `hint` (its DOM paragraph ordinal) which can drift from the XML `w:p`
 * order (docx-preview occasionally emits an extra/fewer `<p>` for breaks etc.).
 * When an `anchor` (the paragraph's current visible text) is given, we trust the
 * TEXT over the number: pick the `w:p` whose combined text equals the anchor,
 * nearest to the hint. Falls back to the hint when there's no anchor or match.
 */
function resolveParagraphIndex(
  paragraphs: Element[],
  hint: number,
  anchor?: string
): number {
  if (anchor === undefined) return hint
  // Exact index already matches the anchor — fast, unambiguous.
  const at = paragraphs[hint]
  if (at && paragraphCombinedText(at) === anchor) return hint
  const matches: number[] = []
  for (let i = 0; i < paragraphs.length; i++) {
    if (paragraphCombinedText(paragraphs[i]) === anchor) matches.push(i)
  }
  if (matches.length === 0) return hint
  matches.sort((a, b) => Math.abs(a - hint) - Math.abs(b - hint))
  return matches[0]
}

/** Combined text of a paragraph's text-bearing runs. */
function paragraphCombinedText(para: Element): string {
  return tag(para, 'w:r')
    .filter((r) => tag(r, 'w:t').length > 0)
    .map(runText)
    .join('')
}

/** Replace a run's text, collapsing it to a single <w:t>. */
function setRunText(run: Element, value: string): void {
  const texts = tag(run, 'w:t')
  texts.forEach((t, i) => {
    if (i === 0) {
      t.textContent = value
      t.setAttribute('xml:space', 'preserve')
    } else {
      t.parentNode?.removeChild(t)
    }
  })
}

function save(zip: PizZip, xml: Document, part: string = DOC_XML): Buffer {
  const out = new XMLSerializer().serializeToString(xml)
  zip.file(part, out)
  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

export interface ParagraphEdit {
  /** The renderer's DOM paragraph ordinal — see resolveParagraphIndex. */
  index: number
  newText: string
  /** The paragraph's text as last known to the caller — see resolveParagraphIndex. */
  anchor?: string
}

/** Applies one already-resolved edit to `paragraphs[idx]`, mutating `xml` in place. Returns whether anything actually changed. */
function applyOneEdit(xml: Document, para: Element, newText: string): boolean {
  const runs = tag(para, 'w:r').filter((r) => tag(r, 'w:t').length > 0)
  const orig = runs.map(runText).join('')
  if (newText === orig) return false
  if (runs.length === 0) {
    // No text runs to carry the edit — append one (copying no formatting).
    const t = xml.createElementNS(W_NS, 'w:t')
    t.setAttribute('xml:space', 'preserve')
    t.appendChild(xml.createTextNode(newText))
    const run = xml.createElementNS(W_NS, 'w:r')
    run.appendChild(t)
    para.appendChild(run)
    return true
  }
  rewriteParagraphRuns(xml, runs, orig, newText)
  return true
}

/**
 * Apply several paragraph text edits in a single zip load/save cycle —
 * the batched form of setParagraphText, for callers that need to rewrite
 * many paragraphs at once (e.g. filling every {{Placeholder}} across a
 * whole document) without reloading and re-deflating the whole zip once
 * per edit, the way calling setParagraphText in a loop would.
 */
export function applyParagraphEdits(buffer: Buffer, edits: ParagraphEdit[], part: string = DOC_XML): Buffer {
  if (edits.length === 0) return buffer
  const { zip, xml } = loadDoc(buffer, part)
  const paragraphs = tag(xml, 'w:p')
  let changed = false
  for (const edit of edits) {
    const idx = resolveParagraphIndex(paragraphs, edit.index, edit.anchor)
    const para = paragraphs[idx]
    if (!para) continue
    if (applyOneEdit(xml, para, edit.newText)) changed = true
  }
  return changed ? save(zip, xml, part) : buffer
}

/**
 * Replace a paragraph's visible text with `newText`, preserving formatting: a
 * longest common prefix/suffix diff keeps every unchanged run (and its rPr)
 * exactly as-is, and only the edited span is rewritten (inheriting the
 * formatting of the run where the edit began). Placeholder runs outside the
 * edited span therefore survive untouched.
 */
export function setParagraphText(
  buffer: Buffer,
  paragraphIndex: number,
  newText: string,
  anchor?: string,
  part: string = DOC_XML
): Buffer {
  return applyParagraphEdits(buffer, [{ index: paragraphIndex, newText, anchor }], part)
}

function normalizeLabel(label: string): string {
  return label.trim().replace(/\s+/g, ' ')
}

/** Every distinct {{Label}} found in the document's body paragraphs, in first-seen order. A placeholder never spans paragraphs. */
export function findPlaceholdersInDocx(buffer: Buffer, part: string = DOC_XML): string[] {
  const { xml } = loadDoc(buffer, part)
  const seen = new Set<string>()
  for (const para of tag(xml, 'w:p')) {
    const text = paragraphCombinedText(para)
    PLACEHOLDER_RE.lastIndex = 0
    let m: RegExpExecArray | null
    while ((m = PLACEHOLDER_RE.exec(text)) !== null) seen.add(normalizeLabel(m[1]))
  }
  return [...seen]
}

/** Every paragraph whose combined text changes under `fillFn`, as ParagraphEdit entries ready for applyParagraphEdits — the anchor is the paragraph's own original text, so a later re-resolve isn't thrown off by an earlier edit elsewhere in the document. */
function computeParagraphEdits(paragraphs: Element[], fillFn: (text: string) => string): ParagraphEdit[] {
  const edits: ParagraphEdit[] = []
  paragraphs.forEach((para, index) => {
    const orig = paragraphCombinedText(para)
    if (!orig.includes('{{')) return
    const filled = fillFn(orig)
    if (filled !== orig) edits.push({ index, newText: filled, anchor: orig })
  })
  return edits
}

function transformPlaceholders(buffer: Buffer, part: string, fillFn: (text: string) => string): Buffer {
  const { zip, xml } = loadDoc(buffer, part)
  const paragraphs = tag(xml, 'w:p')
  const edits = computeParagraphEdits(paragraphs, fillFn)
  let changed = false
  for (const edit of edits) {
    const para = paragraphs[edit.index]
    if (para && applyOneEdit(xml, para, edit.newText)) changed = true
  }
  return changed ? save(zip, xml, part) : buffer
}

/**
 * OOXML-native equivalent of core/createDocument.ts's fillDocumentHtml:
 * replace every {{Label}} occurrence with its resolved value (blank if
 * unresolved), preserving every run's formatting outside the edited spans —
 * see applyParagraphEdits/rewriteParagraphRuns.
 */
export function fillPlaceholdersInDocx(
  buffer: Buffer,
  resolved: PlaceholderMatch[],
  row: Record<string, string>,
  part: string = DOC_XML
): Buffer {
  const valueByLabel = new Map<string, string>()
  for (const r of resolved) valueByLabel.set(r.label, r.column ? row[r.column] ?? '' : '')
  return transformPlaceholders(buffer, part, (text) =>
    text.replace(PLACEHOLDER_RE, (_match, label) => valueByLabel.get(normalizeLabel(label)) ?? '')
  )
}

/**
 * OOXML-native equivalent of core/createDocument.ts's bakeFixedPlaceholders:
 * replace only the given labels' occurrences (matched case-insensitively)
 * with fixed values, leaving every other {{Placeholder}} untouched for later
 * per-row resolution.
 */
export function bakeFixedPlaceholdersInDocx(buffer: Buffer, values: Record<string, string>, part: string = DOC_XML): Buffer {
  const normalized = new Map(Object.entries(values).map(([k, v]) => [normalizeLabel(k).toLowerCase(), v]))
  return transformPlaceholders(buffer, part, (text) =>
    text.replace(PLACEHOLDER_RE, (whole, label) => normalized.get(normalizeLabel(label).toLowerCase()) ?? whole)
  )
}

/**
 * Rewrite a paragraph's run text from `orig` to `newText`, preserving
 * formatting: a longest common prefix/suffix diff keeps every unchanged run
 * (and its rPr) exactly as-is, and only the edited span is rewritten
 * (inheriting the formatting of the run where the edit began). Placeholder runs
 * outside the edited span therefore survive untouched. `runs` must be the
 * paragraph's text-bearing runs and `orig` their concatenated text.
 */
function rewriteParagraphRuns(
  _xml: Document,
  runs: Element[],
  orig: string,
  newText: string
): void {
  if (newText === orig) return

  // Longest common prefix / suffix (in characters).
  let p = 0
  const maxP = Math.min(orig.length, newText.length)
  while (p < maxP && orig[p] === newText[p]) p++
  let s = 0
  while (
    s < Math.min(orig.length - p, newText.length - p) &&
    orig[orig.length - 1 - s] === newText[newText.length - 1 - s]
  )
    s++

  const changeStart = p
  const changeEnd = orig.length - s // exclusive, in orig coordinates
  const newMiddle = newText.slice(p, newText.length - s)

  let cum = 0
  let midPlaced = false
  for (const run of runs) {
    const text = runText(run)
    const rStart = cum
    const rEnd = cum + text.length
    cum = rEnd

    // Run entirely outside the changed region → keep as-is.
    if (rEnd < changeStart || rStart > changeEnd) continue
    if (rEnd === changeStart && rStart !== changeEnd) {
      // Run ends exactly where the change begins. Only relevant for a pure
      // insertion (empty region) — append the new text here to inherit format.
      if (changeStart === changeEnd && !midPlaced) {
        setRunText(run, text + newMiddle)
        midPlaced = true
      }
      continue
    }
    if (rStart === changeEnd && rEnd !== changeStart) {
      continue
    }

    const keepLeft = text.slice(0, Math.max(0, changeStart - rStart))
    const keepRight = text.slice(Math.max(0, changeEnd - rStart))
    const mid = midPlaced ? '' : newMiddle
    midPlaced = true
    setRunText(run, keepLeft + mid + keepRight)
  }

  // Pure insertion that fell past every run (e.g. at the very end).
  if (!midPlaced && newMiddle) {
    const last = runs[runs.length - 1]
    setRunText(last, runText(last) + newMiddle)
  }
}

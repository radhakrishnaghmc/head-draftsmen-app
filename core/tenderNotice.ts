import PizZip from 'pizzip'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import type { Document, Element } from '@xmldom/xmldom'
import { listParagraphs, setParagraphText } from './docx-edit'
import { WORK_TYPE_CATEGORIES, inferTypeOfWork, type WorkTypeCategory } from './workTypeClassifier'

const DOC_XML = 'word/document.xml'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

export interface TenderNoticeWorkItem {
  name: string
  /** Estimate amount, in Lakhs (matching the document's own unit label). */
  amount: string
  /** Works List "Type of work" column value, if the row has one — takes precedence over `inferTypeOfWork(name)` in the page-3 count breakdown (see formatWorkTypeCounts). */
  typeOfWork?: string
}

export interface TenderNoticeGroup {
  /** Download start date, DD.MM.YYYY — the template's default time (2:00 P.M) is left as-is. */
  startDate: string
  /** Download end date, DD.MM.YYYY — also used as the Price Bid Opening date (same day, per the template's default 2:00 P.M / 2:30 P.M times). */
  endDate: string
  /** This group's works, in Sl.No order — continuous across all groups (group 1's items are Sl.No 1..n, group 2's pick up at n+1, etc). */
  items: TenderNoticeWorkItem[]
}

export interface TenderNoticeInput {
  /** New NIT reference number, e.g. "16/DB/EE/Gajularamaram Circle-57/CMC/2026-27". */
  nitNo: string
  /** Today's date, DD.MM.YYYY — replaces the notice's "Dated" fields. */
  today: string
  /**
   * One or more download-date groups. Most notices have exactly one group
   * covering every selected Win Code — a single Events/Schedule block and
   * consolidation row, same as before. A notice recalling works on different
   * timelines needs more than one: each extra group gets its own "Sl No.X to
   * Y" heading, Events/Schedule block and consolidation row.
   */
  groups: TenderNoticeGroup[]
  /**
   * The issuing office's circle/zone (from the chosen Office). The bundled
   * template is a real Gajularamaram Circle-57 / Quthbullapur Zone notice, so
   * every hard-coded place name is swapped to these before saving — without
   * them the notice would still read "Gajularamaram Circle-57" no matter which
   * circle it's issued for. All optional: when absent (e.g. no office chosen),
   * the template's own place names are left untouched.
   */
  circle?: string
  /** Circle number (CNO), e.g. "58" — swapped for the template's "57". */
  circleNumber?: string
  zone?: string
  /** Corporation abbreviation, e.g. "CMC" — swapped for the template's own "CMC" wherever it appears (NIT number, letterhead, eligibility clauses, etc.). */
  corporation?: string
  /** Corporation full name, e.g. "Cyberabad Municipal Corporation" — its upper-cased form is swapped for the letterhead's "CYBERABAD MUNICIPAL CORPORATION". */
  corporationFullName?: string
  /**
   * Contact details shown on the notice (the covering letter's "From" e-mail
   * and the summary table's Tender-Inviting-Authority mobile numbers). The
   * template carries a real past office's values, so these are swapped in when
   * provided; left blank, the template's own values are kept.
   */
  email?: string
  /** Executive Engineer's mobile number. */
  eePhone?: string
  /** Head Draughtsman's mobile number. */
  hdPhone?: string
}

// The bundled template's placeholder values (a real past notice) that get
// replaced wherever they appear, verbatim, across the whole document.
const OLD_NIT = '15/DB/EE/Gajularamaram Circle-57/CMC/2026-27'
// The covering letter's "Lr.No" uses a differently-formatted reference number
// for the same notice — also replaced with the new NIT number.
const OLD_LR_NO = '15/EE/Gajularamaram Circle 57/CMC/2026-27'
// The covering letter's "Ref: Tender Notice No…" is a self-reference to this
// notice, written in yet another abbreviated form (no circle number). Replaced
// with the full new NIT number so the reference matches the notice it's on.
const OLD_REF_NO = '15/DB/EE/Gajularamaram Circle/CMC/2026-27'
const OLD_START_DATE = '16.07.2026'
const OLD_END_DATE = '23.07.2026'
const OLD_DATED = '15.07.2026'
// The template's own office place names (a real Gajularamaram Circle-57 /
// Quthbullapur Zone notice), swapped to the issuing office's circle/zone.
const TPL_CIRCLE_NAME = 'Gajularamaram'
const TPL_CIRCLE_NUMBER = '57'
const TPL_ZONE = 'Quthbullapur'
// The template's own contact details, swapped for the issuing office's when
// supplied. The e-mail also embeds the circle number ("eec57…"), so its swap
// must run before the circle-number swap below (see the swap ordering).
const TPL_EMAIL = 'eec57.GHMC@gmail.com'
const TPL_EE_PHONE = '7893066262'
const TPL_HD_PHONE = '9063836115'
// The template's own corporation branding (NIT number, letterhead,
// eligibility clauses, "Assistant Commissioner PR CMC" line, etc.), swapped
// for the issuing corporation's own abbreviation/full name. Must run AFTER
// the OLD_NIT/OLD_LR_NO/OLD_REF_NO swaps above, since those literal strings
// still contain "CMC" and need to match the template's original text first.
const TPL_CORP = 'CMC'
const TPL_CORP_FULL_CAPS = 'CYBERABAD MUNICIPAL CORPORATION'

/**
 * Pull the circle name and number out of a NIT number like
 * "16/DB/EE/Nizampet Circle-58/CMC/2026-27" → { circle: "Nizampet",
 * circleNumber: "58" }. Used as a fallback so the notice's place names still
 * follow the circle even when the caller didn't pass the office's circle
 * fields explicitly (the NIT number always carries them).
 */
function parseCircleFromNit(nitNo: string): { circle?: string; circleNumber?: string } {
  const m = /EE\/\s*(.+?)\s+Circle[-\s]*(\d+)/i.exec(nitNo)
  if (!m) return {}
  return { circle: m[1].trim(), circleNumber: m[2].trim() }
}

function tag(el: Element | Document, name: string): Element[] {
  return Array.from(el.getElementsByTagName(name)) as unknown as Element[]
}

function paragraphText(p: Element): string {
  return tag(p, 'w:t')
    .map((t) => t.textContent ?? '')
    .join('')
}

function closest(el: Element, tagName: string): Element | null {
  let node: Element | null = el
  while (node && node.tagName !== tagName) node = node.parentNode as Element | null
  return node
}

/** Collapse a cell's first paragraph to a single run carrying `value`, dropping any other runs/paragraphs. */
function setCellText(cell: Element, value: string): void {
  const p = tag(cell, 'w:p')[0]
  if (!p) return
  const runs = tag(p, 'w:r').filter((r) => tag(r, 'w:t').length > 0)
  runs.forEach((run, i) => {
    const texts = tag(run, 'w:t')
    if (i === 0) {
      texts.forEach((t, j) => {
        if (j === 0) {
          t.textContent = value
          t.setAttribute('xml:space', 'preserve')
        } else {
          t.parentNode?.removeChild(t)
        }
      })
    } else {
      run.parentNode?.removeChild(run)
    }
  })
}

/** Sum the items' amounts, formatted to 2 decimals to match the document's own style (e.g. "3.20"). */
function sumAmounts(items: TenderNoticeWorkItem[]): string {
  const total = items.reduce((sum, it) => sum + (Number(String(it.amount).replace(/,/g, '')) || 0), 0)
  return total.toFixed(2)
}

/**
 * Rebuild the itemised work table to have exactly one row per `items` entry,
 * cloning the template's first sample row as the style/formatting reference
 * (Word tables have no formulas to worry about, unlike the Schedule A
 * spreadsheet), and set the Total row to the summed amount.
 */
function fillItemTable(xml: Document, items: TenderNoticeWorkItem[]): void {
  const paragraphs = tag(xml, 'w:p')
  const headerIdx = paragraphs.findIndex((p) => paragraphText(p).trim() === 'Sl.No')
  if (headerIdx === -1) {
    throw new Error('Could not find the "Sl.No" item table header in the Tender Notice template.')
  }
  const headerRow = closest(paragraphs[headerIdx], 'w:tr')
  const table = headerRow && closest(headerRow, 'w:tbl')
  if (!headerRow || !table) {
    throw new Error('Could not find the item table in the Tender Notice template.')
  }

  const allRows = tag(table, 'w:tr')
  const headerRowPos = allRows.indexOf(headerRow)

  const itemRows: Element[] = []
  let totalRow: Element | null = null
  for (let i = headerRowPos + 1; i < allRows.length; i++) {
    const cells = tag(allRows[i], 'w:tc')
    const firstCellText = cells[0] ? tag(cells[0], 'w:p').map(paragraphText).join('').trim() : ''
    if (/^\d+$/.test(firstCellText)) {
      itemRows.push(allRows[i])
    } else if (firstCellText.toLowerCase() === 'total') {
      totalRow = allRows[i]
      break
    }
  }
  if (itemRows.length === 0 || !totalRow) {
    throw new Error('Could not find the item rows / Total row in the Tender Notice template.')
  }

  const referenceRow = itemRows[0]
  for (const row of itemRows) table.removeChild(row)

  for (let i = 0; i < items.length; i++) {
    const clone = referenceRow.cloneNode(true) as Element
    const cells = tag(clone, 'w:tc')
    setCellText(cells[0], String(i + 1))
    setCellText(cells[1], items[i].name)
    setCellText(cells[2], items[i].amount)
    table.insertBefore(clone, totalRow)
  }

  const totalCells = tag(totalRow, 'w:tc')
  setCellText(totalCells[1], sumAmounts(items))
}

/** Find the `<w:tbl>` whose first row's cells open with `headerCells` (e.g. `['Events', 'Schedule']`), content-matched rather than positional since each structural edit above can shift where it sits in the document. */
function findTableByHeader(xml: Document, headerCells: string[]): Element {
  const table = tag(xml, 'w:tbl').find((t) => {
    const firstRow = tag(t, 'w:tr')[0]
    if (!firstRow) return false
    const cells = tag(firstRow, 'w:tc').map((c) => tag(c, 'w:p').map(paragraphText).join('').trim())
    return headerCells.every((h, i) => cells[i] === h)
  })
  if (!table) throw new Error(`Could not find the "${headerCells.join(' / ')}" table in the Tender Notice template.`)
  return table
}

/** A bold, Times New Roman 10pt paragraph — matches the "Sl No.X to Y" group heading style of a real multi-group notice (there's no such heading in the bundled single-group template to clone, so it's built from scratch). */
function makeHeadingParagraph(doc: Document, text: string): Element {
  const rPr = (): Element => {
    const el = doc.createElementNS(W_NS, 'w:rPr') as unknown as Element
    const font = doc.createElementNS(W_NS, 'w:rFonts') as unknown as Element
    font.setAttribute('w:ascii', 'Times New Roman')
    font.setAttribute('w:hAnsi', 'Times New Roman')
    font.setAttribute('w:cs', 'Times New Roman')
    el.appendChild(font)
    el.appendChild(doc.createElementNS(W_NS, 'w:b') as unknown as Element)
    el.appendChild(doc.createElementNS(W_NS, 'w:bCs') as unknown as Element)
    const sz = doc.createElementNS(W_NS, 'w:sz') as unknown as Element
    sz.setAttribute('w:val', '20')
    el.appendChild(sz)
    const szCs = doc.createElementNS(W_NS, 'w:szCs') as unknown as Element
    szCs.setAttribute('w:val', '20')
    el.appendChild(szCs)
    return el
  }
  const p = doc.createElementNS(W_NS, 'w:p') as unknown as Element
  const pPr = doc.createElementNS(W_NS, 'w:pPr') as unknown as Element
  pPr.appendChild(rPr())
  p.appendChild(pPr)
  const r = doc.createElementNS(W_NS, 'w:r') as unknown as Element
  r.appendChild(rPr())
  const t = doc.createElementNS(W_NS, 'w:t') as unknown as Element
  t.setAttribute('xml:space', 'preserve')
  t.textContent = text
  r.appendChild(t)
  p.appendChild(r)
  return p
}

interface GroupRange {
  startSl: number
  endSl: number
  group: TenderNoticeGroup
}

/** Each group's Sl.No range, continuous across all groups (group 1 = Sl.No 1..n, group 2 picks up at n+1, etc — the item table itself is never split, only the Schedule/consolidation blocks below it). */
function computeGroupRanges(groups: TenderNoticeGroup[]): GroupRange[] {
  let sl = 1
  return groups.map((group) => {
    const startSl = sl
    sl += group.items.length
    return { startSl, endSl: sl - 1, group }
  })
}

function slLabel(r: GroupRange): string {
  const pad = (n: number) => String(n).padStart(2, '0')
  return r.startSl === r.endSl ? `Sl No.${pad(r.startSl)}` : `Sl No.${pad(r.startSl)} to ${pad(r.endSl)}`
}

/**
 * Fill the page-1 "Events / Schedule" block: the template carries exactly
 * one (the common single-group case), filled in place. Extra groups each get
 * a cloned block with its own dates, preceded by a bold "Sl No.X to Y"
 * heading — and only then is the first group's block retroactively labelled
 * too, so a single-group notice keeps the template's original heading-free
 * layout untouched.
 */
function fillScheduleGroups(xml: Document, ranges: GroupRange[]): void {
  const table = findTableByHeader(xml, ['Events', 'Schedule'])
  const parent = table.parentNode as unknown as Element

  function setDates(tbl: Element, group: TenderNoticeGroup): void {
    const rows = tag(tbl, 'w:tr')
    const startCell = tag(rows[1], 'w:tc')[1]
    const endCell = tag(rows[2], 'w:tc')[1]
    const priceCell = tag(rows[3], 'w:tc')[1]
    setCellText(startCell, paragraphText(tag(startCell, 'w:p')[0]).replace(OLD_START_DATE, group.startDate))
    setCellText(endCell, paragraphText(tag(endCell, 'w:p')[0]).replace(OLD_END_DATE, group.endDate))
    setCellText(priceCell, paragraphText(tag(priceCell, 'w:p')[0]).replace(OLD_END_DATE, group.endDate))
  }

  // Clone the pristine table before it's mutated below — cloning the
  // already-filled table for later groups would carry group 1's dates into
  // their cells too, leaving nothing left for OLD_START_DATE/OLD_END_DATE to
  // match (a mutate-then-clone ordering bug).
  const pristine = table.cloneNode(true) as Element
  setDates(table, ranges[0].group)
  if (ranges.length === 1) return

  let anchor: Element = table
  for (let i = 1; i < ranges.length; i++) {
    const heading = makeHeadingParagraph(xml, slLabel(ranges[i]))
    const clone = pristine.cloneNode(true) as Element
    setDates(clone, ranges[i].group)
    parent.insertBefore(heading, anchor.nextSibling)
    parent.insertBefore(clone, heading.nextSibling)
    anchor = clone
  }
  parent.insertBefore(makeHeadingParagraph(xml, slLabel(ranges[0])), table)
}

const WORK_TYPE_DISPLAY_LABEL: Record<WorkTypeCategory, string> = {
  FESTIVAL: 'Festival',
  'GRAVE YARDS': 'Graveyard',
  HALLS: 'Halls',
  SPORTS: 'Sports',
  'DESILTING WORKS': 'Desilting',
  'JUNCTION/MEDIAN': 'Junction/Median',
  SWD: 'Storm Water Drainage',
  UGD: 'Under Ground Drainage',
  'PARK/OPEN SPACE': 'Park/Open Space',
  'MAINTENANCE (BT/CC/UGD/OFFICE)': 'Maintenance',
  'BT ROADS': 'BT Roads',
  'CC ROADS': 'CC Roads'
}

// Lets a Works List "Type of work" value that already names a known category
// (by category key or display label, case-insensitively) merge into that
// category's count instead of starting its own one-off bucket.
const CATEGORY_BY_LOWER_NAME = new Map<string, WorkTypeCategory>(
  WORK_TYPE_CATEGORIES.flatMap((c) => [
    [c.toLowerCase(), c] as const,
    [WORK_TYPE_DISPLAY_LABEL[c].toLowerCase(), c] as const
  ])
)

/** The display label this item's row should be counted under — the Works List's own "Type of work" column, if set, takes precedence over inferring one from the work name. */
export function resolveTypeLabel(item: TenderNoticeWorkItem): string {
  const manual = item.typeOfWork?.trim()
  if (manual) {
    const category = CATEGORY_BY_LOWER_NAME.get(manual.toLowerCase())
    return category ? WORK_TYPE_DISPLAY_LABEL[category] : manual
  }
  const category = inferTypeOfWork(item.name) ?? 'MAINTENANCE (BT/CC/UGD/OFFICE)'
  return WORK_TYPE_DISPLAY_LABEL[category]
}

/** "CC Roads: 2, Graveyard: 1, Maintenance: 2" — counts per work-type label (Works List "Type of work" column first, falling back to `inferTypeOfWork` on the name — see resolveTypeLabel). Known categories are listed in `WORK_TYPE_CATEGORIES` order first; any other free-text "Type of work" value follows, in first-seen order. */
function formatWorkTypeCounts(items: TenderNoticeWorkItem[]): string {
  const counts = new Map<string, number>()
  for (const item of items) {
    const label = resolveTypeLabel(item)
    counts.set(label, (counts.get(label) ?? 0) + 1)
  }
  const canonicalLabels = WORK_TYPE_CATEGORIES.map((c) => WORK_TYPE_DISPLAY_LABEL[c])
  const otherLabels = [...counts.keys()].filter((l) => !canonicalLabels.includes(l))
  return [...canonicalLabels.filter((l) => counts.has(l)), ...otherLabels]
    .map((l) => `${l}: ${counts.get(l)}`)
    .join(', ')
}

/**
 * Fill the page-3 consolidation table: same clone-per-extra-group pattern as
 * the Schedule block above, one data row per group (NIT No, a per-work-type
 * count breakdown, summed amount, that group's own download-start/price-bid
 * dates — the Tender-Inviting-Authority contact cell is left untouched in
 * every clone).
 */
function fillConsolidationGroups(xml: Document, nitNo: string, ranges: GroupRange[]): void {
  const table = findTableByHeader(xml, ['NIT No & Date', 'Name of the Work'])
  const parent = table.parentNode as unknown as Element

  function setRow(tbl: Element, r: GroupRange): void {
    const cells = tag(tag(tbl, 'w:tr')[1], 'w:tc')
    setCellText(cells[0], nitNo)
    setCellText(cells[1], formatWorkTypeCounts(r.group.items))
    setCellText(cells[2], sumAmounts(r.group.items))
    setCellText(cells[3], paragraphText(tag(cells[3], 'w:p')[0]).replace(OLD_START_DATE, r.group.startDate))
    setCellText(cells[4], paragraphText(tag(cells[4], 'w:p')[0]).replace(OLD_END_DATE, r.group.endDate))
  }

  // See fillScheduleGroups above for why this must be cloned before `table`
  // is mutated.
  const pristine = table.cloneNode(true) as Element
  setRow(table, ranges[0])
  if (ranges.length === 1) return

  let anchor: Element = table
  for (let i = 1; i < ranges.length; i++) {
    const heading = makeHeadingParagraph(xml, slLabel(ranges[i]))
    const clone = pristine.cloneNode(true) as Element
    setRow(clone, ranges[i])
    parent.insertBefore(heading, anchor.nextSibling)
    parent.insertBefore(clone, heading.nextSibling)
    anchor = clone
  }
  parent.insertBefore(makeHeadingParagraph(xml, slLabel(ranges[0])), table)
}

/**
 * Fill the bundled Tender Notice template: rebuild the itemised work table
 * (one row per selected Win Code across all groups, Total = summed amount),
 * fill the page-1 Schedule block(s) and page-3 consolidation row(s) — one
 * per group, cloned when there's more than one — and swap the NIT number
 * (and the covering letter's differently-formatted Lr.No) and the "Dated"
 * issue date. Times are left untouched, so the template's defaults (2:00
 * P.M download start/end, 2:30 P.M price bid opening) carry through unchanged.
 */
export function fillTenderNotice(buffer: Buffer, input: TenderNoticeInput): Buffer {
  const ranges = computeGroupRanges(input.groups)
  const allItems = input.groups.flatMap((g) => g.items)
  if (allItems.length === 0) throw new Error('Select at least one Win Code.')

  // Step 1 — rebuild the item table, and the Schedule/consolidation blocks
  // (one per group). All structural edits: must run before any
  // paragraph-index-based edits below, since they change how many paragraphs exist.
  const zip = new PizZip(buffer)
  const xmlText = zip.file(DOC_XML)?.asText()
  if (!xmlText) throw new Error('word/document.xml not found in .docx')
  const xml = new DOMParser().parseFromString(xmlText, 'text/xml')
  fillItemTable(xml, allItems)
  fillScheduleGroups(xml, ranges)
  fillConsolidationGroups(xml, input.nitNo, ranges)
  zip.file(DOC_XML, new XMLSerializer().serializeToString(xml))
  let current = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }) as unknown as Buffer

  // Step 2 — NIT/Lr.No/date substitutions, wherever they occur (content-based,
  // so they're unaffected by the item table's now-different paragraph count).
  // Each old→new pair is swapped as its own isolated edit, re-reading the
  // paragraph text before every swap, rather than combining all five pairs
  // into one regex pass over the whole paragraph: the NIT No./Dated line (and
  // the Lr.No/Date line) each carry TWO independent substitutions sharing a
  // paragraph, and diffing the whole paragraph at once treats everything
  // between them — including the unchanged "Dated:" label — as a single
  // changed span, dumping it all into one run and scrambling the two values
  // together (e.g. the new Dated value ending up spliced into the NIT No.).
  // Place-name swaps run AFTER the NIT/Lr.No pairs (which embed the template's
  // circle name+number verbatim) so those match the template text first; only
  // then are the remaining boilerplate mentions of the circle name, circle
  // number and zone rewritten to the issuing office's. Each is skipped when the
  // office value is missing or already equals the template's.
  // Prefer the office's circle fields; when they're absent (e.g. the office
  // wasn't picked and the NIT No was typed by hand), fall back to the circle
  // name+number embedded in the NIT number itself — so the whole notice stays
  // consistent with the circle it's issued for, not just the NIT line.
  const parsed = parseCircleFromNit(input.nitNo)
  const circle = input.circle?.trim() || parsed.circle
  const circleNumber = input.circleNumber?.trim() || parsed.circleNumber

  const placeSwaps: Array<[string, string]> = []
  if (circle && circle !== TPL_CIRCLE_NAME) {
    placeSwaps.push([TPL_CIRCLE_NAME, circle])
  }
  if (circleNumber && circleNumber !== TPL_CIRCLE_NUMBER) {
    // Only within contexts that actually carry the circle number — "Circle-57",
    // "Circle 57", and the EE's e-mail ("eec57…") — never a bare "57", so
    // amounts, phone numbers and other digits are left alone.
    placeSwaps.push([`Circle-${TPL_CIRCLE_NUMBER}`, `Circle-${circleNumber}`])
    placeSwaps.push([`Circle ${TPL_CIRCLE_NUMBER}`, `Circle ${circleNumber}`])
    placeSwaps.push([`eec${TPL_CIRCLE_NUMBER}`, `eec${circleNumber}`])
  }
  if (input.zone && input.zone !== TPL_ZONE) {
    placeSwaps.push([TPL_ZONE, input.zone])
  }

  // Contact details: run BEFORE placeSwaps so the full template e-mail
  // ("eec57.GHMC@gmail.com") is matched before the circle-number swap would
  // rewrite its "eec57" fragment. Each is skipped when not supplied.
  const contactSwaps: Array<[string, string]> = []
  if (input.email?.trim()) contactSwaps.push([TPL_EMAIL, input.email.trim()])
  if (input.eePhone?.trim()) contactSwaps.push([TPL_EE_PHONE, input.eePhone.trim()])
  if (input.hdPhone?.trim()) contactSwaps.push([TPL_HD_PHONE, input.hdPhone.trim()])

  const corpSwaps: Array<[string, string]> = []
  if (input.corporation?.trim() && input.corporation.trim() !== TPL_CORP) {
    corpSwaps.push([TPL_CORP, input.corporation.trim()])
  }
  if (input.corporationFullName?.trim()) {
    const fullCaps = input.corporationFullName.trim().toUpperCase()
    if (fullCaps !== TPL_CORP_FULL_CAPS) corpSwaps.push([TPL_CORP_FULL_CAPS, fullCaps])
  }

  const swaps: Array<[string, string]> = [
    [OLD_NIT, input.nitNo],
    [OLD_LR_NO, input.nitNo],
    [OLD_REF_NO, input.nitNo],
    [OLD_DATED, input.today],
    ...contactSwaps,
    ...placeSwaps,
    ...corpSwaps
  ]
  for (const [oldStr, newStr] of swaps) {
    if (oldStr === newStr) continue
    let paragraphs = listParagraphs(current)
    for (let i = 0; i < paragraphs.length; i++) {
      if (!paragraphs[i].includes(oldStr)) continue
      const next = paragraphs[i].split(oldStr).join(newStr)
      current = setParagraphText(current, i, next, paragraphs[i])
      paragraphs = listParagraphs(current)
    }
  }

  return current
}

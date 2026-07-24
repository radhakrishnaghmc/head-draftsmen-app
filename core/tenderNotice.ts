import PizZip from 'pizzip'
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'
import type { Document, Element } from '@xmldom/xmldom'
import { listParagraphs, setParagraphText } from './docx-edit'

const DOC_XML = 'word/document.xml'

export interface TenderNoticeWorkItem {
  name: string
  /** Estimate amount, in Lakhs (matching the document's own unit label). */
  amount: string
}

export interface TenderNoticeInput {
  /** New NIT reference number, e.g. "16/DB/EE/Gajularamaram Circle-57/CMC/2026-27". */
  nitNo: string
  /** Download start date, DD.MM.YYYY — the template's default time (2:00 P.M) is left as-is. */
  startDate: string
  /** Download end date, DD.MM.YYYY — also used as the Price Bid Opening date (same day, per the template's default 2:00 P.M / 2:30 P.M times). */
  endDate: string
  /** Today's date, DD.MM.YYYY — replaces the notice's "Dated" fields. */
  today: string
  /** One or more works (from selected Win Codes) — one item-table row each. */
  items: TenderNoticeWorkItem[]
}

// The bundled template's placeholder values (a real past notice) that get
// replaced wherever they appear, verbatim, across the whole document.
const OLD_NIT = '15/DB/EE/Gajularamaram Circle-57/CMC/2026-27'
// The covering letter's "Lr.No" uses a differently-formatted reference number
// for the same notice — also replaced with the new NIT number.
const OLD_LR_NO = '15/EE/Gajularamaram Circle 57/CMC/2026-27'
const OLD_START_DATE = '16.07.2026'
const OLD_END_DATE = '23.07.2026'
const OLD_DATED = '15.07.2026'
// The bottom summary table's sample work name — located by content search
// since the itemised table above it can grow/shrink the paragraph indices.
const OLD_SUMMARY_NAME = 'Electrical works(3 Nos) '

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

/**
 * Fill the bundled Tender Notice template: rebuild the itemised work table
 * (one row per selected Win Code, Total = summed amount), swap the NIT number
 * (and the covering letter's differently-formatted Lr.No), the three dates
 * (download start/end, and the "Dated" issue date), and the bottom summary
 * table's work count ("No of works: N")/amount. Times are left untouched, so the template's
 * defaults (2:00 P.M download start/end, 2:30 P.M price bid opening) carry
 * through unchanged.
 */
export function fillTenderNotice(buffer: Buffer, input: TenderNoticeInput): Buffer {
  if (input.items.length === 0) throw new Error('Select at least one Win Code.')

  // Step 1 — rebuild the item table (structural edit: must run before any
  // paragraph-index-based edits, since it changes how many paragraphs exist).
  const zip = new PizZip(buffer)
  const xmlText = zip.file(DOC_XML)?.asText()
  if (!xmlText) throw new Error('word/document.xml not found in .docx')
  const xml = new DOMParser().parseFromString(xmlText, 'text/xml')
  fillItemTable(xml, input.items)
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
  const swaps: Array<[string, string]> = [
    [OLD_NIT, input.nitNo],
    [OLD_LR_NO, input.nitNo],
    [OLD_START_DATE, input.startDate],
    [OLD_END_DATE, input.endDate],
    [OLD_DATED, input.today]
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

  // Step 3 — bottom summary table's work name/amount, located by content
  // search (its paragraph indices shifted by however many rows step 1 added
  // or removed above it).
  const refreshed = listParagraphs(current)
  const summaryNameIdx = refreshed.findIndex((t) => t === OLD_SUMMARY_NAME)
  if (summaryNameIdx !== -1) {
    const workCount = `No of works: ${input.items.length}`
    current = setParagraphText(current, summaryNameIdx, workCount, refreshed[summaryNameIdx])
    const amountIdx = summaryNameIdx + 1
    const afterName = listParagraphs(current)
    current = setParagraphText(current, amountIdx, sumAmounts(input.items), afterName[amountIdx])
  }

  return current
}

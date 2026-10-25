import { PDFDocument, StandardFonts, rgb, type PDFFont } from 'pdf-lib'
import type { MonitoringFormatRow, MonitoringFormatSummary } from './monitoringFormat'
import { indianDigitGroups } from './worksAmounts'

// US Legal landscape (14 x 8.5 in), in points.
const PAGE_W = 1008
const PAGE_H = 612
const MARGIN = 24
let FONT_SIZE = 8
let LINE_H = 10.5
const PAD_X = 4
const PAD_Y = 4
/** A date / status cell longer than this is treated as free text and allowed to wrap. */
const ONE_LINE_MAX = 12

/** The built-in PDF fonts only cover WinAnsi — swap anything else for '?' so drawing never throws. */
function clean(text: string, font: PDFFont): string {
  let out = ''
  for (const ch of (text ?? '').replace(/\s+/g, ' ').trim()) {
    try {
      font.encodeText(ch)
      out += ch
    } catch {
      out += '?'
    }
  }
  return out
}

/** Greedy word wrap to a pixel width; words longer than the column are hard-split. */
function wrap(text: string, font: PDFFont, width: number): string[] {
  if (!text) return ['']
  const lines: string[] = []
  let line = ''
  const push = (w: string) => {
    const candidate = line ? `${line} ${w}` : w
    if (font.widthOfTextAtSize(candidate, FONT_SIZE) <= width) {
      line = candidate
      return
    }
    if (line) lines.push(line)
    line = ''
    let rest = w
    while (font.widthOfTextAtSize(rest, FONT_SIZE) > width && rest.length > 1) {
      // Prefer breaking a too-wide word after a hyphen ("274-" / "Bachupally") over cutting it mid-letter.
      let n = 0
      for (let k = rest.indexOf('-') + 1; k > 0 && k < rest.length; k = rest.indexOf('-', k) + 1) {
        if (font.widthOfTextAtSize(rest.slice(0, k), FONT_SIZE) <= width) n = k
        else break
      }
      if (n === 0) {
        n = rest.length - 1
        while (n > 1 && font.widthOfTextAtSize(rest.slice(0, n), FONT_SIZE) > width) n--
      }
      lines.push(rest.slice(0, n))
      rest = rest.slice(n)
    }
    line = rest
  }
  text.split(' ').forEach(push)
  if (line) lines.push(line)
  return lines.length ? lines : ['']
}

/** Landscape Legal table: every row grows to fit its full wrapped text, the header repeats on each page. Columns are whatever the caller passes (the workbook's own), widths sized from their content. */
export async function buildWorksListPdf(opts: WorksListPdfOpts): Promise<Uint8Array> {
  if (opts.fontSize) return (await renderWorksList(opts)).bytes
  // Largest font (11 pt down to 6.5 pt) at which every full page holds at least MIN_WORKS_PER_PAGE works.
  const want = Math.min(MIN_WORKS_PER_PAGE, opts.rows.length)
  let out = await renderWorksList({ ...opts, fontSize: 11 })
  for (let size = 10.5; size >= 6.5 && !(Math.min(...out.rowsPerPage.slice(0, -1), Infinity) >= want || out.rowsPerPage.length <= 1); size -= 0.5) {
    out = await renderWorksList({ ...opts, fontSize: size })
  }
  return out.bytes
}

/** Each works-list PDF page should carry at least this many works; the font shrinks until it does. */
const MIN_WORKS_PER_PAGE = 9

interface WorksListPdfOpts {
  title: string
  subtitle?: string
  headers: string[]
  rows: string[][]
  /** Overrides the column-count based default font size (pt). */
  fontSize?: number
  /** Row indexes drawn bold on a shaded band (totals) … */
  boldRows?: number[]
  /** … and row indexes that are a section heading: only the first cell, shaded, spanning the table. */
  sectionRows?: number[]
  /** … and row indexes that are an empty spacer line: no text, fill or rules. */
  gapRows?: number[]
}

async function renderWorksList(opts: WorksListPdfOpts): Promise<{ bytes: Uint8Array; pageCount: number; overflowsWidth: boolean; rowsPerPage: number[] }> {
  const pdf = await PDFDocument.create()
  const font = await pdf.embedFont(StandardFonts.Helvetica)
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold)

  FONT_SIZE = opts.fontSize ?? (opts.headers.length > 14 ? 8 : opts.headers.length > 11 ? 9 : 11)
  LINE_H = FONT_SIZE + 2.5

  const COLUMNS = opts.headers.map((title, i) => {
    const vals = opts.rows.map((r) => (r[i] ?? '').trim())
    const filled = vals.filter(Boolean)
    const numeric = filled.length > 0 && filled.every((v) => /^[\d,.\s-]+$/.test(v))
    const avg = filled.length ? filled.reduce((s, v) => s + v.length, 0) / filled.length : 0
    const longestWord = Math.max(...title.split(/\s+/).map((w) => w.length), ...filled.map((v) => Math.max(...v.split(/\s+/).map((w) => w.length))).slice(0, 200))
    // weight ~ typical text length, never narrower than the longest unbreakable word (capped)
    const weight = Math.max(Math.min(avg, 90), Math.min(longestWord, 16), 4) * (/name.*work/i.test(title) ? 2.8 : /remark/i.test(title) ? 1.6 : 1)
    return { title, key: i, weight, right: numeric, noWrap: /date|status/i.test(title) }
  })

  const tableW = PAGE_W - MARGIN * 2
  const totalWeight = COLUMNS.reduce((s, c) => s + c.weight, 0)
  const widths = COLUMNS.map((c) => (c.weight / totalWeight) * tableW)
  // A header such as "Estimate (Rs. Lakhs)" should wrap between words, never mid-word: give every
  // column at least the width of its longest heading word, taking the extra from the roomier columns.
  // Date and status columns stay on one line (headings too, unless a heading is long like "Target date of
  // completion"), and no column is ever narrower than its longest unbreakable word, so a word is never split
  // mid-way ("Progres/s").
  const headNoWrap = (c: (typeof COLUMNS)[number]) => c.noWrap && c.title.length <= 9
  const minWidths = COLUMNS.map((c, i) => {
    const headW = headNoWrap(c)
      ? bold.widthOfTextAtSize(clean(c.title, bold), FONT_SIZE)
      : Math.max(...c.title.split(/\s+/).map((w) => bold.widthOfTextAtSize(clean(w, bold), FONT_SIZE)))
    const cellW = opts.rows.map((r) => {
      const t = clean(r[i] ?? '', font)
      // Only real dates / statuses are kept on one line; a stray long note typed into such a column wraps
      // like normal text instead of stretching the column and squeezing "Name of Work".
      if (c.noWrap && t.length <= ONE_LINE_MAX) return font.widthOfTextAtSize(t, FONT_SIZE)
      return Math.min(72, Math.max(0, ...t.split(' ').map((w) => font.widthOfTextAtSize(w, FONT_SIZE))))
    })
    return Math.max(headW, ...cellW) + PAD_X * 2 + 2
  })
  for (let pass = 0; pass < 4; pass++) {
    const short = widths.reduce((sum, w, i) => sum + Math.max(0, minWidths[i] - w), 0)
    if (short < 0.01) break
    const spare = widths.reduce((sum, w, i) => sum + Math.max(0, w - minWidths[i]), 0)
    if (spare <= 0) break
    widths.forEach((w, i) => {
      widths[i] = w < minWidths[i] ? minWidths[i] : w - (Math.max(0, w - minWidths[i]) / spare) * Math.min(short, spare)
    })
  }
  // If even the minimums don't fit (very many columns), scale everything down so the table never runs off the page.
  const sumW = widths.reduce((a, w) => a + w, 0)
  if (sumW > tableW + 0.5) widths.forEach((w, i) => (widths[i] = (w / sumW) * tableW))
  const xs = widths.reduce<number[]>((acc, w, i) => [...acc, (acc[i] ?? MARGIN) + w], [MARGIN]).slice(0, -1)

  let page = pdf.addPage([PAGE_W, PAGE_H])
  let y = PAGE_H - MARGIN
  const pages = [page]
  const rowsPerPage = [0]

  // Light vertical rules between the columns (and at both table edges), drawn per row so they follow the row heights.
  const drawColumnLines = (top: number, h: number) => {
    const edges = [...xs, MARGIN + tableW]
    edges.forEach((x) => page.drawLine({ start: { x, y: top }, end: { x, y: top - h }, thickness: 0.4, color: rgb(0.85, 0.85, 0.85) }))
  }
  const drawHeader = () => {
    const lines = COLUMNS.map((c, i) => (headNoWrap(c) ? [clean(c.title, bold)] : wrap(clean(c.title, bold), bold, widths[i] - PAD_X * 2)))
    const h = Math.max(...lines.map((l) => l.length)) * LINE_H + PAD_Y * 2
    page.drawRectangle({ x: MARGIN, y: y - h, width: tableW, height: h, color: rgb(0.94, 0.96, 0.98) })
    COLUMNS.forEach((c, i) => {
      lines[i].forEach((ln, li) => {
        const w = bold.widthOfTextAtSize(ln, FONT_SIZE)
        const x = c.right ? xs[i] + widths[i] - PAD_X - w : xs[i] + PAD_X
        page.drawText(ln, { x, y: y - PAD_Y - FONT_SIZE - li * LINE_H, size: FONT_SIZE, font: bold, color: rgb(0.1, 0.1, 0.1) })
      })
    })
    drawColumnLines(y, h)
    page.drawLine({ start: { x: MARGIN, y: y - h }, end: { x: MARGIN + tableW, y: y - h }, thickness: 0.8, color: rgb(0.5, 0.5, 0.5) })
    y -= h
  }

  page.drawText(clean(opts.title, bold), { x: MARGIN, y: y - 14, size: 14, font: bold, color: rgb(0.1, 0.1, 0.1) })
  y -= 22
  if (opts.subtitle) {
    page.drawText(clean(opts.subtitle, font), { x: MARGIN, y: y - 9, size: 9, font, color: rgb(0.35, 0.35, 0.35) })
    y -= 16
  }
  y -= 6
  drawHeader()

  const boldSet = new Set(opts.boldRows ?? [])
  const sectionSet = new Set(opts.sectionRows ?? [])
  const gapSet = new Set(opts.gapRows ?? [])
  for (const [ri, row] of opts.rows.entries()) {
    const isSection = sectionSet.has(ri)
    if (gapSet.has(ri)) {
      y -= LINE_H + PAD_Y * 2
      continue
    }
    const rowFont = isSection || boldSet.has(ri) ? bold : font
    const cells = isSection
      ? [[clean(row[0] ?? '', bold)], ...COLUMNS.slice(1).map(() => [''])]
      : COLUMNS.map((c, i) => (c.noWrap && clean(row[i] ?? '', rowFont).length <= ONE_LINE_MAX ? [clean(row[i] ?? '', rowFont)] : wrap(clean(row[i] ?? '', rowFont), rowFont, widths[i] - PAD_X * 2)))
    const h = Math.max(...cells.map((l) => l.length)) * LINE_H + PAD_Y * 2
    if (y - h < MARGIN + 14) {
      page = pdf.addPage([PAGE_W, PAGE_H])
      pages.push(page)
      rowsPerPage.push(0)
      y = PAGE_H - MARGIN
      drawHeader()
    }
    rowsPerPage[rowsPerPage.length - 1]++
    if (isSection || boldSet.has(ri)) {
      page.drawRectangle({ x: MARGIN, y: y - h, width: tableW, height: h, color: isSection ? rgb(0.88, 0.92, 0.97) : rgb(0.96, 0.96, 0.96) })
    }
    COLUMNS.forEach((c, i) => {
      cells[i].forEach((ln, li) => {
        const w = rowFont.widthOfTextAtSize(ln, FONT_SIZE)
        const x = c.right && !isSection ? xs[i] + widths[i] - PAD_X - w : xs[i] + PAD_X
        page.drawText(ln, { x, y: y - PAD_Y - FONT_SIZE - li * LINE_H, size: FONT_SIZE, font: rowFont, color: rgb(0.12, 0.12, 0.12) })
      })
    })
    if (!isSection) drawColumnLines(y, h)
    page.drawLine({ start: { x: MARGIN, y: y - h }, end: { x: MARGIN + tableW, y: y - h }, thickness: 0.4, color: rgb(0.8, 0.8, 0.8) })
    y -= h
  }

  pages.forEach((p, i) => {
    const label = `Page ${i + 1} of ${pages.length}`
    p.drawText(label, {
      x: PAGE_W - MARGIN - font.widthOfTextAtSize(label, 8),
      y: MARGIN - 10,
      size: 8,
      font,
      color: rgb(0.4, 0.4, 0.4)
    })
  })
  return { bytes: await pdf.save(), pageCount: pages.length, rowsPerPage, overflowsWidth: widths.reduce((a, w) => a + w, 0) > tableW + 0.5 }
}

// ---- Monitoring Format "Abstract": drawn to look like the online sheet (Times, black grid, merged headings) ----

/** Excel column widths of the sheet's Abstract (A = item type / ward, then No. + Amt pairs); the PDF keeps the same proportions. */
export const ABSTRACT_COL_UNITS = [34, 9, 13, ...Array<number>(20).fill(9)]

export const ABSTRACT_GROUP_KEYS: (keyof Omit<MonitoringFormatRow, 'itemType'>)[] = [
  'totalWorks',
  'completed',
  'upto25',
  'upto50',
  'upto75',
  'above75',
  'progressTotal',
  'toBeStarted',
  'tenderProcess',
  'heldUp',
  'cancelled'
]

/** A number the way the sheet shows it: no thousands separators, at most two decimals, no trailing zeros. */
export function sheetNumber(n: number): string {
  return String(Math.round((Number.isFinite(n) ? n : 0) * 100) / 100)
}

/**
 * The Monitoring Format Abstract on one Legal-landscape sheet, laid out like the online sheet: a title bar,
 * the "Name of the Circle / Date" line (values in red), the three-row merged heading (Item type | Total works
 * | Completed | Progress (upto 25% … Above 75%, TOTAL) | To be Started | Tender Process | Held Up | Cancelled,
 * each with No. and Amt in Rs Lakhs), the item-type rows and Total, then a blank gap and the same heading
 * over the ward-wise (or circle-wise) rows. The font steps down until everything fits the page.
 */
export async function buildAbstractPdf(summary: MonitoringFormatSummary, sheetTitle: string): Promise<Uint8Array> {
  const circleLevel = summary.wardPivotKind !== 'circle'
  const sizes = [11, 10.5, 10, 9.5, 9, 8.5, 8, 7.5, 7, 6.5, 6]
  let last: { bytes: Uint8Array; fits: boolean } | null = null
  for (const fs of sizes) {
    last = await renderAbstract(summary, sheetTitle, circleLevel, fs)
    if (last.fits) break
  }
  return last!.bytes
}

async function renderAbstract(
  summary: MonitoringFormatSummary,
  sheetTitle: string,
  circleLevel: boolean,
  fs: number
): Promise<{ bytes: Uint8Array; fits: boolean }> {
  const pdf = await PDFDocument.create()
  const reg = await pdf.embedFont(StandardFonts.TimesRoman)
  const bld = await pdf.embedFont(StandardFonts.TimesRomanBold)
  const page = pdf.addPage([PAGE_W, PAGE_H])
  const M = 18
  const tableW = PAGE_W - M * 2
  const unitSum = ABSTRACT_COL_UNITS.reduce((a, b) => a + b, 0)
  const colW = ABSTRACT_COL_UNITS.map((u) => (u / unitSum) * tableW)
  const colX = colW.reduce<number[]>((acc, w, i) => [...acc, acc[i] + w], [M]).slice(0, -1)
  const black = rgb(0, 0, 0)
  const red = rgb(1, 0, 0)
  const LH = 1.15
  let fits = true
  let y = PAGE_H - M

  const lines = (text: string, font: PDFFont, size: number, width: number): string[] => {
    const out: string[] = []
    let line = ''
    for (const w of clean(text, font).split(' ')) {
      const cand = line ? `${line} ${w}` : w
      if (font.widthOfTextAtSize(cand, size) <= width || !line) line = cand
      else {
        out.push(line)
        line = w
      }
    }
    if (line) out.push(line)
    return out.length ? out : ['']
  }
  /** One bordered (possibly merged) cell spanning columns c0..c1, drawn with its top at the current y. */
  const cell = (
    c0: number,
    c1: number,
    top: number,
    h: number,
    text: string,
    o: { bold?: boolean; size?: number; align?: 'left' | 'center' | 'right'; color?: ReturnType<typeof rgb>; border?: boolean | 'tb' } = {}
  ) => {
    const x = colX[c0]
    const w = colX[c1] + colW[c1] - x
    if (o.border === 'tb') {
      page.drawLine({ start: { x, y: top }, end: { x: x + w, y: top }, thickness: 0.6, color: black })
      page.drawLine({ start: { x, y: top - h }, end: { x: x + w, y: top - h }, thickness: 0.6, color: black })
    } else if (o.border !== false) {
      page.drawRectangle({ x, y: top - h, width: w, height: h, borderColor: black, borderWidth: 0.6 })
    }
    if (!text) return
    const font = o.bold ? bld : reg
    const size = o.size ?? fs
    const ls = lines(text, font, size, w - 4)
    const textH = ls.length * size * LH
    ls.forEach((ln, i) => {
      const tw = font.widthOfTextAtSize(ln, size)
      const tx = o.align === 'left' ? x + 3 : o.align === 'right' ? x + w - 3 - tw : x + (w - tw) / 2
      page.drawText(ln, { x: tx, y: top - (h - textH) / 2 - size * 0.95 - i * size * LH, size, font, color: o.color ?? black })
    })
  }
  const need = (h: number) => {
    if (y - h < M) fits = false
  }

  // Title bar and the circle / date line
  const titleH = fs * 2.6
  need(titleH)
  cell(0, 22, y, titleH, sheetTitle, { bold: true, size: fs * 1.25 })
  y -= titleH
  const infoH = fs * 2.2
  need(infoH)
  cell(0, 22, y, infoH, '', { border: 'tb' })
  cell(1, 3, y, infoH, circleLevel ? 'Name of the Circle:-' : 'Name of the Zone:-', { bold: true, size: fs * 1.15, border: false })
  cell(4, 7, y, infoH, summary.officeLabel, { bold: true, size: fs * 1.15, color: red, border: false })
  cell(18, 19, y, infoH, 'Date:-', { bold: true, size: fs * 1.15, border: false })
  cell(20, 22, y, infoH, summary.asOfDate ?? '', { bold: true, size: fs * 1.15, color: red, border: false })
  y -= infoH

  const table = (firstHeading: string, rows: MonitoringFormatRow[], totals: MonitoringFormatRow | undefined, itemRows: boolean) => {
    // Three heading rows (the sheet merges them: row 1-2 for most groups, row 3 holds No. / Amt)
    const h1 = fs * 1.7
    const h2 = fs * 1.7
    const h3 = fs * LH * 3 + 3
    need(h1 + h2 + h3)
    cell(0, 0, y, h1 + h2 + h3, firstHeading, { bold: true })
    const groupHead: [number, number, string][] = [
      [1, 2, 'Total works'],
      [3, 4, 'Completed'],
      [15, 16, 'To be Started'],
      [17, 18, 'Tender Process'],
      [19, 20, 'Held Up'],
      [21, 22, 'Cancelled']
    ]
    groupHead.forEach(([c0, c1, t]) => cell(c0, c1, y, h1 + h2, t, { bold: true }))
    cell(5, 12, y, h1, 'Progress', { bold: true })
    cell(13, 14, y, h1, '', {})
    ;(
      [
        [5, 6, 'upto 25%'],
        [7, 8, 'upto 50%'],
        [9, 10, 'upto 75%'],
        [11, 12, 'Above 75%'],
        [13, 14, 'TOTAL']
      ] as [number, number, string][]
    ).forEach(([c0, c1, t]) => cell(c0, c1, y - h1, h2, t, { bold: true }))
    for (let c = 1; c <= 22; c += 2) {
      cell(c, c, y - h1 - h2, h3, 'No.', { bold: true })
      cell(c + 1, c + 1, y - h1 - h2, h3, 'Amt in Rs Lakhs', { bold: true })
    }
    y -= h1 + h2 + h3

    const body = [...rows.map((r) => ({ r, name: r.itemType, total: false })), ...(totals ? [{ r: totals, name: 'Total', total: true }] : [])]
    for (const { r, name, total } of body) {
      const nameSize = itemRows ? fs : fs
      const nameLines = lines(name, total || itemRows ? bld : reg, nameSize, colW[0] - 6)
      const h = Math.max(fs * 1.7, nameLines.length * nameSize * LH + 4)
      need(h)
      cell(0, 0, y, h, name, { bold: total, align: 'left', size: nameSize })
      ABSTRACT_GROUP_KEYS.forEach((k, g) => {
        ;[r[k].no, r[k].amt].forEach((v, j) => {
          const c = 1 + g * 2 + j
          const text = sheetNumber(v)
          if ((total ? bld : reg).widthOfTextAtSize(text, fs) > colW[c] - 4) fits = false
          cell(c, c, y, h, text, { bold: total, align: 'right' })
        })
      })
      y -= h
    }
  }

  table('Item type', summary.rows, summary.totals, true)
  if (summary.wardRows && summary.wardRows.length > 0) {
    y -= fs * 2 // one clear blank line between the two tables
    table(summary.wardPivotKind === 'circle' ? 'Circle Name and number' : 'Ward Name and number', summary.wardRows, summary.wardTotals, false)
  }
  return { bytes: await pdf.save(), fits }
}

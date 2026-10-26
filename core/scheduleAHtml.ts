import * as ExcelJS from 'exceljs'

/**
 * A filled Schedule-A workbook drawn as printable HTML, for the PDF download.
 *
 * The workbook's amounts are live formulas with no stored results (they are
 * computed when Excel opens it), so they are evaluated here. Everything inside
 * the print area is drawn with the sheet's own fonts, borders, merges and
 * widths. The item table's header prints once, on the first page only, and the Total / Tender
 * Quoted / signature rows are kept together so the Executive Engineer block is
 * never stranded alone or cut off.
 */

const LAST_COL = 6 // A..F — the print area; G/H hold scratch formulas that must not print

interface Region {
  top: number
  left: number
  bottom: number
  right: number
}

export interface ScheduleAHtml {
  html: string
  /** Natural width of the table in CSS px (before fit-to-page zoom). */
  tableWidthPx: number
  /** The sheet's own page margins, in inches. */
  margins: { top: number; bottom: number; left: number; right: number }
}

function colNumber(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

function parseRegion(range: string): Region | null {
  const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(range.trim())
  if (!m) return null
  return { left: colNumber(m[1]), top: Number(m[2]), right: colNumber(m[3]), bottom: Number(m[4]) }
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function argbToCss(argb: string | undefined): string | null {
  if (!argb || argb.length < 6) return null
  return `#${argb.slice(-6)}`
}

// ---- formula evaluation -----------------------------------------------------

type CellGetter = (row: number, col: number) => number

/** Evaluate the small formula dialect Schedule A uses: + - * / ( ) postfix %, cell refs, SUM(range), ROUND(x,n). */
function evaluateFormula(formula: string, get: CellGetter): number {
  const src = formula.replace(/^=/, '')
  let pos = 0
  const peek = (): string => {
    while (src[pos] === ' ') pos++
    return src[pos] ?? ''
  }
  const refAt = (): { col: number; row: number } | null => {
    const m = /^\$?([A-Z]+)\$?(\d+)/i.exec(src.slice(pos))
    if (!m) return null
    pos += m[0].length
    return { col: colNumber(m[1]), row: Number(m[2]) }
  }

  const parseExpr = (): number => {
    let v = parseTerm()
    for (;;) {
      const c = peek()
      if (c === '+') {
        pos++
        v += parseTerm()
      } else if (c === '-') {
        pos++
        v -= parseTerm()
      } else return v
    }
  }
  const parseTerm = (): number => {
    let v = parseUnary()
    for (;;) {
      const c = peek()
      if (c === '*') {
        pos++
        v *= parseUnary()
      } else if (c === '/') {
        pos++
        v /= parseUnary()
      } else return v
    }
  }
  const parseUnary = (): number => {
    const c = peek()
    if (c === '+') {
      pos++
      return parseUnary()
    }
    if (c === '-') {
      pos++
      return -parseUnary()
    }
    return parsePostfix()
  }
  const parsePostfix = (): number => {
    let v = parsePrimary()
    while (peek() === '%') {
      pos++
      v /= 100
    }
    return v
  }
  const parseRange = (): number[] => {
    const a = refAt()
    if (!a) return []
    if (peek() !== ':') return [get(a.row, a.col)]
    pos++
    const b = refAt()
    if (!b) return [get(a.row, a.col)]
    const out: number[] = []
    for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++)
      for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) out.push(get(r, c))
    return out
  }
  const parsePrimary = (): number => {
    const c = peek()
    if (c === '(') {
      pos++
      const v = parseExpr()
      if (peek() === ')') pos++
      return v
    }
    const num = /^\d+(\.\d+)?/.exec(src.slice(pos))
    if (num) {
      pos += num[0].length
      return Number(num[0])
    }
    const fn = /^([A-Z]+)\(/i.exec(src.slice(pos))
    if (fn) {
      pos += fn[0].length
      const name = fn[1].toUpperCase()
      let result = NaN
      if (name === 'SUM') {
        result = 0
        for (;;) {
          for (const n of parseRange()) result += n
          if (peek() === ',') pos++
          else break
        }
      } else if (name === 'ROUND') {
        const x = parseExpr()
        let digits = 0
        if (peek() === ',') {
          pos++
          digits = parseExpr()
        }
        const f = Math.pow(10, digits)
        result = Math.round((x + Number.EPSILON) * f) / f
      }
      if (peek() === ')') pos++
      return result
    }
    const ref = refAt()
    if (ref) return get(ref.row, ref.col)
    return NaN
  }

  const v = parseExpr()
  return Number.isFinite(v) ? v : NaN
}

function formatNumber(n: number, numFmt: string | undefined): string {
  if (!Number.isFinite(n)) return ''
  const fmt = numFmt && numFmt !== 'General' ? numFmt : ''
  if (!fmt) return String(Number(n.toFixed(10)))
  const percent = fmt.includes('%')
  const decimals = (/\.([0#]+)/.exec(fmt)?.[1] ?? '').length
  return (
    new Intl.NumberFormat('en-US', {
      useGrouping: fmt.includes(','),
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    }).format(percent ? n * 100 : n) + (percent ? '%' : '')
  )
}

// ---- styling ----------------------------------------------------------------

function borderCss(side: string, b: Partial<ExcelJS.Border> | undefined): string {
  if (!b || !b.style) return ''
  const width =
    b.style === 'medium' || b.style === 'mediumDashed' ? 2 : b.style === 'thick' || b.style === 'double' ? 3 : 1
  const kind =
    b.style === 'double'
      ? 'double'
      : b.style === 'hair' || b.style === 'dotted'
        ? 'dotted'
        : /dash/i.test(b.style)
          ? 'dashed'
          : 'solid'
  const color = argbToCss((b.color as { argb?: string } | undefined)?.argb) ?? '#000'
  return `border-${side}:${width}px ${kind} ${color};`
}

function fontCss(f: Partial<ExcelJS.Font> | undefined): string {
  const css: string[] = []
  css.push(`font-family:'${f?.name ?? 'Arial'}',Arial,Helvetica,sans-serif;`) // single quotes: this lands in a style="..." attribute
  css.push(`font-size:${f?.size ?? 10}pt;`)
  if (f?.bold) css.push('font-weight:bold;')
  if (f?.italic) css.push('font-style:italic;')
  if (f?.underline) css.push('text-decoration:underline;')
  const color = argbToCss((f?.color as { argb?: string } | undefined)?.argb)
  if (color) css.push(`color:${color};`)
  return css.join('')
}

// ---- document ---------------------------------------------------------------

export async function scheduleAWorkbookToHtml(xlsx: Buffer): Promise<ScheduleAHtml> {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(xlsx as unknown as ArrayBuffer)
  const ws = wb.worksheets[0]
  if (!ws) throw new Error('Schedule A workbook has no sheet.')

  // Never stop short of the signature block, whatever print area the file carries.
  let lastRow = ws.rowCount
  while (lastRow > 1 && !rowHasContent(ws, lastRow)) lastRow--

  const regions = ((ws.model as { merges?: string[] }).merges ?? [])
    .map(parseRegion)
    .filter((r): r is Region => !!r && r.top <= lastRow && r.left <= LAST_COL)
    .map((r) => ({ ...r, right: Math.min(r.right, LAST_COL), bottom: Math.min(r.bottom, lastRow) }))
  const regionAt = (row: number, col: number): Region | undefined =>
    regions.find((r) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right)

  // Formula results, evaluated on demand (and memoised).
  const memo = new Map<string, number>()
  const numAt: CellGetter = (row, col) => {
    const key = `${row},${col}`
    const hit = memo.get(key)
    if (hit !== undefined) return hit
    memo.set(key, 0) // guards a circular reference
    const cell = ws.getCell(row, col)
    let n = 0
    if (cell.type === ExcelJS.ValueType.Formula) n = evaluateFormula(String(cell.formula ?? ''), numAt)
    else if (typeof cell.value === 'number') n = cell.value
    else if (typeof cell.value === 'string') n = Number(cell.value.replace(/,/g, '')) || 0
    memo.set(key, n)
    return n
  }

  const widths = Array.from({ length: LAST_COL }, (_, i) => Math.round((ws.getColumn(i + 1).width ?? 9) * 7 + 5))
  const tableWidthPx = widths.reduce((a, b) => a + b, 0)

  const cellHtml = (row: number, col: number): string => {
    const region = regionAt(row, col)
    if (region && (region.top !== row || region.left !== col)) return '' // swallowed by a merge
    const cell = ws.getCell(row, col)
    const top = region?.top ?? row
    const left = region?.left ?? col
    const bottom = region?.bottom ?? row
    const right = region?.right ?? col
    const b = {
      top: ws.getCell(top, left).border?.top,
      left: ws.getCell(top, left).border?.left,
      bottom: ws.getCell(bottom, left).border?.bottom,
      right: ws.getCell(top, right).border?.right
    }
    const align = cell.alignment ?? {}
    const isNumber = cell.type === ExcelJS.ValueType.Formula || typeof cell.value === 'number'
    const horizontal =
      align.horizontal === 'centerContinuous'
        ? 'center'
        : align.horizontal && align.horizontal !== 'fill'
          ? align.horizontal
          : isNumber
            ? 'right'
            : 'left'
    const vertical = align.vertical === 'middle' ? 'middle' : align.vertical === 'top' ? 'top' : 'bottom'
    const fill =
      cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid'
        ? argbToCss((cell.fill.fgColor as { argb?: string } | undefined)?.argb)
        : null

    let inner = ''
    if (cell.type === ExcelJS.ValueType.Formula) {
      inner = esc(formatNumber(numAt(row, col), cell.numFmt))
    } else if (cell.value && typeof cell.value === 'object' && 'richText' in cell.value) {
      inner = cell.value.richText
        .map(
          (run) =>
            `<span style="${run.font ? fontCss({ ...cell.font, ...run.font } as Partial<ExcelJS.Font>) : ''}">${esc(run.text)}</span>`
        )
        .join('')
    } else if (typeof cell.value === 'number') {
      inner = esc(formatNumber(cell.value, cell.numFmt))
    } else if (cell.value && typeof cell.value === 'object' && 'text' in cell.value) {
      inner = esc(String((cell.value as { text: unknown }).text ?? ''))
    } else if (cell.value != null && typeof cell.value !== 'object') {
      inner = esc(String(cell.value))
    }

    const style =
      fontCss(cell.font) +
      `text-align:${horizontal};vertical-align:${vertical};` +
      `white-space:${align.wrapText ? 'pre-wrap' : 'pre'};` +
      (align.indent ? `padding-left:${3 + align.indent * 9}px;` : '') +
      (fill ? `background:${fill};` : '') +
      borderCss('top', b.top) +
      borderCss('left', b.left) +
      borderCss('bottom', b.bottom) +
      borderCss('right', b.right)
    const span = `${region && region.right > region.left ? ` colspan="${region.right - region.left + 1}"` : ''}${
      region && region.bottom > region.top ? ` rowspan="${region.bottom - region.top + 1}"` : ''
    }`
    return `<td${span} style="${style}">${inner}</td>`
  }

  const rowHtml = (row: number): string => {
    const height = Math.round(((ws.getRow(row).height ?? 15) * 4) / 3)
    let tds = ''
    for (let c = 1; c <= LAST_COL; c++) tds += cellHtml(row, c)
    return `<tr style="height:${height}px">${tds}</tr>`
  }
  const rowsHtml = (from: number, to: number): string => {
    let out = ''
    for (let r = from; r <= to; r++) out += rowHtml(r)
    return out
  }
  const colgroup = `<colgroup>${widths.map((w) => `<col style="width:${w}px">`).join('')}</colgroup>`
  const table = (head: string, body: string): string =>
    `<table style="width:${tableWidthPx}px">${colgroup}${head ? `<thead>${head}</thead>` : ''}<tbody>${body}</tbody></table>`

  // Where the sheet splits: preamble | item table (header repeats) | total, tender quoted, signature.
  let headerRow = 1
  for (let r = 1; r <= Math.min(lastRow, 40); r++) {
    if (/^item\s*no\.?$/i.test(String(ws.getCell(r, 1).text ?? '').trim())) {
      headerRow = r
      break
    }
  }
  let tailStart = lastRow + 1
  for (let r = headerRow + 2; r <= lastRow; r++) {
    const f = ws.getCell(r, LAST_COL)
    if (f.type === ExcelJS.ValueType.Formula && /^SUM\(/i.test(String(f.formula ?? ''))) {
      tailStart = r
      break
    }
  }
  const hasSplit = headerRow > 1 && tailStart <= lastRow

  const body = hasSplit
    ? table('', rowsHtml(1, headerRow - 1)) +
      table('', rowsHtml(headerRow, tailStart - 1)) +
      `<div class="keep">${table('', rowsHtml(tailStart, lastRow))}</div>`
    : table('', rowsHtml(1, lastRow))

  const m = ws.pageSetup.margins
  const margins = { top: m?.top ?? 0.55, bottom: m?.bottom ?? 0.43, left: m?.left ?? 0.38, right: m?.right ?? 0.24 }

  const html = `<html><head><meta charset="utf-8"><style>
  html, body { margin: 0; padding: 0; background: #fff }
  table { border-collapse: collapse; table-layout: fixed }
  td { box-sizing: border-box; padding: 1px 3px; overflow: hidden; word-wrap: break-word }
  tr { break-inside: avoid }
  .keep { break-inside: avoid }
</style></head><body><div id="sheet">${body}</div></body></html>`

  return { html, tableWidthPx, margins }
}

function rowHasContent(ws: ExcelJS.Worksheet, row: number): boolean {
  for (let c = 1; c <= LAST_COL; c++) {
    const cell = ws.getCell(row, c)
    if (cell.type === ExcelJS.ValueType.Formula) return true
    if (String(cell.text ?? '').trim() !== '') return true
  }
  return false
}

import * as ExcelJS from 'exceljs'

export type PrintPaper = 'a4' | 'a3' | 'letter' | 'legal'
export type PrintOrientation = 'auto' | 'portrait' | 'landscape'

export interface ExcelPrintOptions {
  /** Sheet names to include, in workbook order; empty/omitted = every visible sheet with content. */
  sheets?: string[]
  paper?: PrintPaper
  orientation?: PrintOrientation
}

export interface ExcelPrintSheetInfo {
  name: string
  rows: number
  cols: number
  /** Natural width of the printed range in CSS px. */
  widthPx: number
  /** Orientation "auto" would pick for this sheet. */
  orientation: 'portrait' | 'landscape'
}

// Paper sizes in inches, portrait [width, height].
const PAPER_IN: Record<PrintPaper, [number, number]> = {
  a4: [8.27, 11.69],
  a3: [11.69, 16.54],
  letter: [8.5, 11],
  legal: [8.5, 14]
}
const PAPER_CSS: Record<PrintPaper, string> = { a4: 'A4', a3: 'A3', letter: 'Letter', legal: 'Legal' }

// Margins applied to every page: a bit more top/bottom than sides so the page-number footer has room.
const MARGIN_SIDE_IN = 0.5
const MARGIN_TOP_IN = 0.55
const MARGIN_BOTTOM_IN = 0.7
// A sheet is turned landscape rather than shrunk below this fraction of its natural size.
const MIN_PORTRAIT_SCALE = 0.72
// Never shrink below this (unreadable); wider sheets are clipped at the page edge instead.
const MIN_SCALE = 0.35
const DEFAULT_COL_PX = 64
const DEFAULT_ROW_PX = 20

interface Range {
  top: number
  left: number
  bottom: number
  right: number
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

const colPx = (col: ExcelJS.Column | undefined): number => {
  if (col?.hidden) return 0
  return col?.width ? Math.round(col.width * 7 + 5) : DEFAULT_COL_PX
}

const rowPx = (row: ExcelJS.Row | undefined): number => {
  if (row?.hidden) return 0
  return row?.height ? Math.round((row.height * 4) / 3) : DEFAULT_ROW_PX
}

function parseA1(ref: string): { row: number; col: number } {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref.trim())
  if (!m) return { row: 1, col: 1 }
  let col = 0
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { row: Number(m[2]), col }
}

function parseRange(ref: string): Range | null {
  const [a, b] = ref.replace(/^.*!/, '').split(':')
  if (!a) return null
  const s = parseA1(a)
  const e = b ? parseA1(b) : s
  return {
    top: Math.min(s.row, e.row),
    left: Math.min(s.col, e.col),
    bottom: Math.max(s.row, e.row),
    right: Math.max(s.col, e.col)
  }
}

function cellHasContent(cell: ExcelJS.Cell): boolean {
  if (cell.value !== null && cell.value !== undefined && cell.text !== '') return true
  const b = cell.border
  if (b && (b.top || b.bottom || b.left || b.right)) return true
  const f = cell.fill
  return !!(f && f.type === 'pattern' && f.pattern === 'solid')
}

/** The sheet's own print area when it defines one, otherwise the used range trimmed of empty edge rows/columns. */
function printedRange(ws: ExcelJS.Worksheet): Range | null {
  const area = ws.pageSetup?.printArea
  if (area) {
    const r = parseRange(area.split(',')[0])
    if (r) return r
  }
  let top = Infinity
  let left = Infinity
  let bottom = 0
  let right = 0
  ws.eachRow({ includeEmpty: false }, (row, rn) => {
    row.eachCell({ includeEmpty: true }, (cell, cn) => {
      if (!cellHasContent(cell)) return
      top = Math.min(top, rn)
      bottom = Math.max(bottom, rn)
      left = Math.min(left, cn)
      right = Math.max(right, cn)
    })
  })
  if (!isFinite(top) || bottom === 0) return null
  // Merged blocks can extend past the last filled cell.
  for (const m of ws.model.merges ?? []) {
    const r = parseRange(m)
    if (r && r.top >= top && r.left >= left && r.top <= bottom && r.left <= right) {
      bottom = Math.max(bottom, r.bottom)
      right = Math.max(right, r.right)
    }
  }
  return { top, left, bottom, right }
}

function argb(c: Partial<ExcelJS.Color> | undefined): string | null {
  if (!c?.argb) return null
  const v = c.argb.length === 8 ? c.argb.slice(2) : c.argb
  return /^[0-9a-fA-F]{6}$/.test(v) ? `#${v}` : null
}

const BORDER_PX: Record<string, string> = {
  thin: '1px solid',
  hair: '1px dotted',
  dotted: '1px dotted',
  dashed: '1px dashed',
  dashDot: '1px dashed',
  dashDotDot: '1px dashed',
  mediumDashed: '2px dashed',
  medium: '2px solid',
  double: '3px double',
  thick: '3px solid'
}

function borderCss(side: string, b: Partial<ExcelJS.Border> | undefined): string {
  if (!b?.style) return ''
  return `border-${side}:${BORDER_PX[b.style] ?? '1px solid'} ${argb(b.color) ?? '#000'};`
}

const pad2 = (n: number) => String(n).padStart(2, '0')

/** Display text for a cell: numbers follow their number format's decimals / thousands / percent, dates dd-mm-yyyy. */
function displayText(cell: ExcelJS.Cell): string {
  let v: unknown = cell.value
  if (v && typeof v === 'object' && !(v instanceof Date)) {
    const o = v as { result?: unknown; richText?: { text: string }[]; text?: string; error?: string }
    if (o.richText) return o.richText.map((r) => r.text).join('')
    if (o.error) return o.error
    if ('result' in o) v = o.result
    else if (o.text !== undefined) return String(o.text)
  }
  if (v === null || v === undefined) return ''
  if (v instanceof Date) return `${pad2(v.getUTCDate())}-${pad2(v.getUTCMonth() + 1)}-${v.getUTCFullYear()}`
  if (typeof v === 'number') {
    const fmt = (cell.numFmt ?? '').split(';')[0]
    if (!fmt || fmt === 'General') return Number.isInteger(v) ? String(v) : String(Math.round(v * 1e10) / 1e10)
    const pct = fmt.includes('%')
    const dec = /\.([0#?]+)/.exec(fmt)?.[1].length ?? 0
    const n = pct ? v * 100 : v
    const s = fmt.includes(',')
      ? n.toLocaleString('en-IN', { minimumFractionDigits: dec, maximumFractionDigits: dec })
      : n.toFixed(dec)
    const currency = /[₹$€£]/.exec(fmt)?.[0] ?? ''
    return currency + s + (pct ? '%' : '')
  }
  return String(v)
}

function cellCss(cell: ExcelJS.Cell, isNumber: boolean): string {
  const css: string[] = []
  const f = cell.font
  if (f) {
    if (f.bold) css.push('font-weight:700')
    if (f.italic) css.push('font-style:italic')
    const deco = [f.underline ? 'underline' : '', f.strike ? 'line-through' : ''].filter(Boolean).join(' ')
    if (deco) css.push(`text-decoration:${deco}`)
    if (f.size) css.push(`font-size:${Math.round(((f.size * 4) / 3) * 10) / 10}px`)
    if (f.name) css.push(`font-family:'${f.name.replace(/'/g, '')}',Calibri,Carlito,Arial,Helvetica,sans-serif`)
    const c = argb(f.color)
    if (c) css.push(`color:${c}`)
  }
  const fill = cell.fill
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const c = argb(fill.fgColor)
    if (c) css.push(`background:${c}`)
  }
  const al = cell.alignment
  const h = al?.horizontal
  css.push(`text-align:${h === 'centerContinuous' ? 'center' : h && h !== 'fill' ? h : isNumber ? 'right' : 'left'}`)
  css.push(`vertical-align:${al?.vertical === 'top' ? 'top' : al?.vertical === 'middle' ? 'middle' : 'bottom'}`)
  if (al?.indent) css.push(`padding-left:${3 + al.indent * 9}px`)
  css.push(al?.wrapText ? 'white-space:pre-wrap;word-break:break-word' : 'white-space:nowrap;overflow:hidden')
  const b = cell.border
  if (b)
    css.push(
      borderCss('top', b.top) + borderCss('bottom', b.bottom) + borderCss('left', b.left) + borderCss('right', b.right)
    )
  return css.join(';')
}

/** The print-title rows (repeated on every page) when the sheet sets them, as [first,last]. */
function titleRows(ws: ExcelJS.Worksheet, range: Range): [number, number] | null {
  const t = ws.pageSetup?.printTitlesRow
  const m = t ? /\$?(\d+):\$?(\d+)/.exec(t) : null
  if (!m) return null
  const a = Number(m[1])
  const b = Number(m[2])
  return a >= range.top && b <= range.bottom && a <= b ? [a, b] : null
}

interface SheetHtml {
  info: ExcelPrintSheetInfo
  table: string
  tableWidthPx: number
}

function sheetToHtml(ws: ExcelJS.Worksheet): SheetHtml | null {
  const range = printedRange(ws)
  if (!range) return null

  // Merge map: master cell → span, covered cell → skipped.
  const masters = new Map<string, { rs: number; cs: number }>()
  const covered = new Set<string>()
  for (const m of ws.model.merges ?? []) {
    const r = parseRange(m)
    if (!r) continue
    masters.set(`${r.top},${r.left}`, { rs: r.bottom - r.top + 1, cs: r.right - r.left + 1 })
    for (let rr = r.top; rr <= r.bottom; rr++)
      for (let cc = r.left; cc <= r.right; cc++) if (rr !== r.top || cc !== r.left) covered.add(`${rr},${cc}`)
  }

  const widths: number[] = []
  for (let c = range.left; c <= range.right; c++) widths.push(colPx(ws.getColumn(c)))
  const tableWidthPx = widths.reduce((a, b) => a + b, 0)
  if (tableWidthPx === 0) return null
  const titles = titleRows(ws, range)

  const rowHtml = (rn: number): string => {
    const row = ws.getRow(rn)
    const h = rowPx(row)
    if (h === 0) return ''
    const cells: string[] = []
    for (let c = range.left; c <= range.right; c++) {
      if (widths[c - range.left] === 0 || covered.has(`${rn},${c}`)) continue
      const cell = row.getCell(c)
      const span = masters.get(`${rn},${c}`)
      // Spans are clipped to the printed range.
      const cs = span ? Math.min(span.cs, range.right - c + 1) : 1
      const rs = span ? Math.min(span.rs, range.bottom - rn + 1) : 1
      const attrs = (cs > 1 ? ` colspan="${cs}"` : '') + (rs > 1 ? ` rowspan="${rs}"` : '')
      cells.push(`<td${attrs} style="${cellCss(cell, typeof cell.value === 'number')}">${esc(displayText(cell))}</td>`)
    }
    return `<tr style="height:${h}px">${cells.join('')}</tr>`
  }

  const head: string[] = []
  const body: string[] = []
  for (let rn = range.top; rn <= range.bottom; rn++) {
    const html = rowHtml(rn)
    if (titles && rn >= titles[0] && rn <= titles[1]) head.push(html)
    else body.push(html)
  }
  const cols = widths.map((w) => (w === 0 ? '' : `<col style="width:${w}px">`)).join('')
  const table =
    `<table style="width:${tableWidthPx}px"><colgroup>${cols}</colgroup>` +
    (head.length ? `<thead>${head.join('')}</thead>` : '') +
    `<tbody>${body.join('')}</tbody></table>`

  return {
    info: {
      name: ws.name,
      rows: range.bottom - range.top + 1,
      cols: widths.filter((w) => w > 0).length,
      widthPx: tableWidthPx,
      orientation: 'portrait'
    },
    table,
    tableWidthPx
  }
}

async function loadWorkbook(xlsx: Buffer | Uint8Array): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(Buffer.from(xlsx) as unknown as ExcelJS.Buffer)
  return wb
}

function pickOrientation(
  tableWidthPx: number,
  paper: PrintPaper,
  requested: PrintOrientation
): 'portrait' | 'landscape' {
  if (requested !== 'auto') return requested
  const portraitPx = (PAPER_IN[paper][0] - 2 * MARGIN_SIDE_IN) * 96
  return portraitPx / tableWidthPx >= MIN_PORTRAIT_SCALE ? 'portrait' : 'landscape'
}

/** Sheets with something to print, with the orientation Auto would choose — for the picker list. */
export async function inspectExcelForPrint(
  xlsx: Buffer | Uint8Array,
  paper: PrintPaper = 'a4'
): Promise<ExcelPrintSheetInfo[]> {
  const wb = await loadWorkbook(xlsx)
  const out: ExcelPrintSheetInfo[] = []
  for (const ws of wb.worksheets) {
    if (ws.state !== 'visible') continue
    const s = sheetToHtml(ws)
    if (s) out.push({ ...s.info, orientation: pickOrientation(s.tableWidthPx, paper, 'auto') })
  }
  return out
}

/**
 * One self-contained HTML document with every chosen sheet as its own run of pages. Each sheet gets a named CSS page
 * (size + orientation) so portrait and landscape sheets can sit in the same PDF; the table is scaled down to the
 * printable width, margins are even, and any print area / repeated title rows the workbook defines are honoured.
 */
export async function excelToPrintHtml(
  xlsx: Buffer | Uint8Array,
  opts: ExcelPrintOptions = {}
): Promise<{ html: string; sheets: number }> {
  const paper = opts.paper ?? 'a4'
  const requested = opts.orientation ?? 'auto'
  const wb = await loadWorkbook(xlsx)
  const wanted = opts.sheets && opts.sheets.length ? new Set(opts.sheets) : null

  const pageRules: string[] = []
  const sections: string[] = []
  let n = 0
  for (const ws of wb.worksheets) {
    if (ws.state !== 'visible' || (wanted && !wanted.has(ws.name))) continue
    const s = sheetToHtml(ws)
    if (!s) continue
    const orient = pickOrientation(s.tableWidthPx, paper, requested)
    const [pw, ph] = PAPER_IN[paper]
    const printablePx = ((orient === 'portrait' ? pw : ph) - 2 * MARGIN_SIDE_IN) * 96
    const zoom = Math.max(MIN_SCALE, Math.min(1, printablePx / s.tableWidthPx))
    const page = `p${n}`
    pageRules.push(
      `@page ${page} { size: ${PAPER_CSS[paper]} ${orient}; margin: ${MARGIN_TOP_IN}in ${MARGIN_SIDE_IN}in ${MARGIN_BOTTOM_IN}in ${MARGIN_SIDE_IN}in }`
    )
    sections.push(
      `<section style="page:${page}${n ? ';break-before:page' : ''}"><div style="zoom:${zoom}">${s.table}</div></section>`
    )
    n++
  }
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
${pageRules.join('\n')}
html, body { margin: 0; padding: 0; background: #fff; font-family: Calibri, Carlito, Arial, Helvetica, sans-serif; font-size: 14.67px; color: #000 }
table { border-collapse: collapse; table-layout: fixed }
td { padding: 1px 3px; line-height: 1.25; box-sizing: border-box }
tr { break-inside: avoid }
thead { display: table-header-group }
</style></head><body>${sections.join('')}</body></html>`
  return { html, sheets: n }
}

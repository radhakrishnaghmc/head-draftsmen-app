import * as ExcelJS from 'exceljs'
import type { MbModel } from './measurementBook'
import { SEIG_CHART, SEIG_DMFT_SHARE, SEIG_PERMIT_SHARE, SEIG_RATES, SEIG_SMET_SHARE } from './seigniorageChart'

// Styling copied from the reference sheet measurement_N-271-2025-26.xlsx: Calibri 12, thin borders, green headings,
// pale-green item rows, pale-yellow totals, blue = typed-in value, black = formula.
const GREEN = '1E5631'
const FILL_TITLE = 'A9D08E'
const FILL_HEAD = 'C6E0B4'
const FILL_ITEM = 'E2EFDA'
const FILL_TOTAL = 'FFF2CC'
const BLUE = '0000FF'
const BLACK = '000000'
const THIN = { style: 'thin' as const, color: { argb: 'FF000000' } }
const BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN }

type Al = Partial<ExcelJS.Alignment>
const CENTER: Al = { horizontal: 'center', vertical: 'middle' }
const RIGHT: Al = { horizontal: 'right', vertical: 'middle' }
const TOPWRAP: Al = { vertical: 'top', wrapText: true }
const MID_WRAP: Al = { vertical: 'middle', wrapText: true }

function style(
  c: ExcelJS.Cell,
  o: { bold?: boolean; color?: string; fill?: string; al?: Al; nf?: string; size?: number } = {}
) {
  c.font = { name: 'Calibri', size: o.size ?? 12, bold: !!o.bold, color: { argb: `FF${o.color ?? BLACK}` } }
  if (o.fill) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${o.fill}` } }
  c.border = BOX
  if (o.al) c.alignment = o.al
  if (o.nf) c.numFmt = o.nf
}

function setup(ws: ExcelJS.Worksheet, lastCol: number, widths: number[], margin: number, titleRows?: string) {
  widths.forEach((w, i) => (ws.getColumn(i + 1).width = w))
  ws.views = [{ showGridLines: false }]
  ws.pageSetup = {
    orientation: 'landscape',
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: margin, right: margin, top: 1, bottom: 1, header: 0.3, footer: 0.3 },
    ...(titleRows ? { printTitlesRow: titleRows } : {})
  }
  void lastCol
}

function title(ws: ExcelJS.Worksheet, text: string, lastCol: number, size: number, height: number) {
  ws.mergeCells(1, 1, 1, lastCol)
  for (let c = 1; c <= lastCol; c++) style(ws.getCell(1, c), { bold: true, color: GREEN, fill: FILL_TITLE, size })
  const c = ws.getCell(1, 1)
  c.value = text
  c.alignment = { horizontal: 'left', vertical: 'middle' }
  ws.getRow(1).height = height
}

const LABELS = [
  'Name of work',
  'Name of agency / contractor',
  'Estimate amount',
  'Contract amount',
  'Tender percentage',
  'Date of mark-out',
  'Date of completion'
]

function headerBlock(ws: ExcelJS.Worksheet, m: MbModel, lastCol: number, labelSpan: number) {
  const h = m.header
  const vals = [h.work, h.agency, h.estimate, h.contract, h.percentage, h.markOut, h.completion]
  LABELS.forEach((label, i) => {
    const r = 2 + i
    if (labelSpan > 1) ws.mergeCells(r, 1, r, labelSpan)
    ws.mergeCells(r, labelSpan + 1, r, lastCol)
    for (let c = 1; c <= labelSpan; c++)
      style(ws.getCell(r, c), { bold: true, color: GREEN, fill: FILL_HEAD, al: MID_WRAP })
    for (let c = labelSpan + 1; c <= lastCol; c++) style(ws.getCell(r, c), { al: MID_WRAP })
    ws.getCell(r, 1).value = label
    ws.getCell(r, labelSpan + 1).value = vals[i] || null
    ws.getRow(r).height = 38
  })
  ws.getRow(11).height = 16
}

function headRow(ws: ExcelJS.Worksheet, row: number, heads: string[]) {
  heads.forEach((t, i) => {
    const c = ws.getCell(row, i + 1)
    c.value = t
    style(c, {
      bold: true,
      color: GREEN,
      fill: FILL_HEAD,
      al: { horizontal: 'center', vertical: 'middle', wrapText: true }
    })
  })
  ws.getRow(row).height = 40
}

/** Row height for wrapped text in a column of `chars` characters. */
const wrapHeight = (text: string, chars: number, min = 40) => Math.max(min, Math.ceil(text.length / chars) * 21 + 6)

/** Measurements + Abstract + Seigniorage charges workbook, with live formulas, laid out like the reference measurement sheet. */
export async function buildMeasurementBookWorkbook(m: MbModel, mbName: string): Promise<Buffer> {
  const wb = new ExcelJS.Workbook()

  // ---------------- Measurements ----------------
  const ms = wb.addWorksheet('Measurements')
  setup(ms, 17, [7, 9, 60, 7, 10, 9, 9, 9, 10, 9, 9, 9, 10, 10, 15, 15, 10], 0.4, '13:13')
  title(ms, `MEASUREMENT BOOK ${mbName} | MEASUREMENTS`, 17, 15, 32)
  headerBlock(ms, m, 17, 2)
  headRow(ms, 13, [
    'Sl',
    'PDF page',
    'Description / location',
    'No.',
    'L',
    'B1',
    'B2',
    'B3',
    'B (avg)',
    'D1',
    'D2',
    'D3',
    'D (avg)',
    'Factor',
    'Contents (calc)',
    'Contents (as written)',
    'Diff'
  ])

  const totalRow: number[] = []
  let r = 14
  m.items.forEach((it, idx) => {
    // item heading row
    for (let c = 1; c <= 17; c++) style(ms.getCell(r, c), { bold: true, fill: FILL_ITEM, al: c >= 4 ? RIGHT : CENTER })
    style(ms.getCell(r, 2), { bold: true, color: BLUE, fill: FILL_ITEM, al: CENTER })
    style(ms.getCell(r, 3), { bold: true, fill: FILL_ITEM, al: TOPWRAP })
    ms.getCell(r, 2).value = it.mbPage
    ms.getCell(r, 3).value = it.description
    ms.getRow(r).height = wrapHeight(it.description, 62)
    r++

    const first = r
    const contd = new Map(it.contd.map((c) => [c.atLine, c.page]))
    let n = 0
    it.lines.forEach((ln, i) => {
      const page = contd.get(i)
      if (page != null) {
        for (let c = 1; c <= 17; c++) style(ms.getCell(r, c), { al: c >= 4 ? RIGHT : CENTER })
        style(ms.getCell(r, 2), { color: BLUE, al: CENTER })
        style(ms.getCell(r, 3), { al: TOPWRAP })
        ms.getCell(r, 2).value = page
        const d = it.description.length > 90 ? `${it.description.slice(0, 90)}...` : it.description
        ms.getCell(r, 3).value = `(contd.) ${d}`
        ms.getRow(r).height = 40
        r++
      }
      n++
      for (let c = 1; c <= 17; c++) style(ms.getCell(r, c), { al: RIGHT, nf: '0.00' })
      style(ms.getCell(r, 1), { al: CENTER })
      style(ms.getCell(r, 2), { al: CENTER })
      style(ms.getCell(r, 3), { al: TOPWRAP })
      style(ms.getCell(r, 4), { color: BLUE, al: RIGHT })
      ;[5, 6, 7, 8, 10, 11, 12, 16].forEach((c) => style(ms.getCell(r, c), { color: BLUE, al: RIGHT, nf: '0.00' }))
      ;[9, 13].forEach((c) => style(ms.getCell(r, c), { al: RIGHT, nf: '0.000' }))
      style(ms.getCell(r, 14), { color: BLUE, al: RIGHT, nf: '0.0000' })
      ms.getCell(r, 1).value = n
      ms.getCell(r, 3).value = ln.location || null
      if (ln.no != null) ms.getCell(r, 4).value = ln.no
      if (ln.l != null) ms.getCell(r, 5).value = ln.l
      ln.b.forEach((v, k) => v != null && (ms.getCell(r, 6 + k).value = v))
      ln.d.forEach((v, k) => v != null && (ms.getCell(r, 10 + k).value = v))
      if (ln.b.some((v) => v != null)) ms.getCell(r, 9).value = { formula: `AVERAGE(F${r}:H${r})` }
      if (ln.d.some((v) => v != null)) ms.getCell(r, 13).value = { formula: `AVERAGE(J${r}:L${r})` }
      ms.getCell(r, 15).value = {
        formula: `ROUND(D${r}*IF(E${r}="",1,E${r})*IF(I${r}="",1,I${r})*IF(M${r}="",1,M${r})*IF(N${r}="",1,N${r}),2)`
      }
      if (ln.written != null) {
        ms.getCell(r, 16).value = ln.written
        ms.getCell(r, 17).value = { formula: `ROUND(O${r}-P${r},2)` }
      }
      r++
    })

    // item total row
    for (let c = 1; c <= 17; c++)
      style(ms.getCell(r, c), { bold: true, fill: FILL_TOTAL, al: c >= 4 ? RIGHT : CENTER, nf: '0.00' })
    style(ms.getCell(r, 3), { bold: true, fill: FILL_TOTAL, al: MID_WRAP })
    style(ms.getCell(r, 16), { bold: true, color: BLUE, fill: FILL_TOTAL, al: RIGHT, nf: '0.00' })
    ms.getCell(r, 3).value = `Item ${idx + 1} total (as written${it.unit ? `, ${it.unit}` : ''})`
    if (it.lines.length) ms.getCell(r, 15).value = { formula: `SUM(O${first}:O${r - 1})` }
    const written =
      it.totalWritten ??
      (it.lines.some((l) => l.written != null) ? it.lines.reduce((a, l) => a + (l.written ?? 0), 0) : null)
    if (written != null) {
      ms.getCell(r, 16).value = written
      ms.getCell(r, 17).value = { formula: `ROUND(O${r}-P${r},2)` }
    }
    totalRow.push(r)
    r++
  })

  // ---------------- Abstract ----------------
  const ab = wb.addWorksheet('Abstract')
  setup(ab, 12, [17, 64, 12, 14, 14, 10, 9, 13, 9, 16, 16, 11], 0.4, '13:13')
  title(ab, `MEASUREMENT BOOK ${mbName} | ABSTRACT OF QUANTITIES AND AMOUNTS`, 12, 15, 32)
  headerBlock(ab, m, 12, 1)
  headRow(ab, 13, [
    'Item',
    'Description',
    'Qty 1 (from Measurements)',
    'Total qty (calc)',
    'Total qty (as written)',
    'Qty diff',
    'Unit',
    'Rate (Rs)',
    'Per',
    'Amount (calc)',
    'Amount (as written)',
    'Diff'
  ])
  m.abstract.forEach((a, i) => {
    const row = 14 + i
    const item = m.items[i]
    for (let c = 1; c <= 12; c++) style(ab.getCell(row, c), { al: RIGHT, nf: '0.00' })
    style(ab.getCell(row, 1), { al: CENTER, nf: 'General' })
    style(ab.getCell(row, 2), { al: TOPWRAP, nf: 'General' })
    style(ab.getCell(row, 5), { color: BLUE, al: RIGHT, nf: '0.00' })
    style(ab.getCell(row, 7), { color: BLUE, al: CENTER, nf: 'General' })
    style(ab.getCell(row, 8), { color: BLUE, al: RIGHT, nf: '0.00' })
    style(ab.getCell(row, 9), { color: BLUE, al: CENTER, nf: 'General' })
    ;[10, 12].forEach((c) => style(ab.getCell(row, c), { al: RIGHT, nf: '#,##0.00' }))
    style(ab.getCell(row, 11), { color: BLUE, al: RIGHT, nf: '#,##0.00' })
    ab.getCell(row, 1).value = String(i + 1)
    ab.getCell(row, 2).value = item ? `${a.description} - Qty vide pg. no. ${item.mbPage} MB` : a.description
    if (totalRow[i]) ab.getCell(row, 3).value = { formula: `Measurements!O${totalRow[i]}` }
    ab.getCell(row, 4).value = { formula: `ROUND(SUM(C${row}:C${row}),2)` }
    if (a.qtyWritten != null) {
      ab.getCell(row, 5).value = a.qtyWritten
      ab.getCell(row, 6).value = { formula: `ROUND(D${row}-E${row},2)` }
    }
    ab.getCell(row, 7).value = a.unit || null
    if (a.rate != null) {
      ab.getCell(row, 8).value = a.rate
      ab.getCell(row, 9).value = `1 ${a.unit}`.trim()
      ab.getCell(row, 10).value = { formula: `ROUND(D${row}*H${row},2)` }
    }
    if (a.amountWritten != null) {
      ab.getCell(row, 11).value = a.amountWritten
      ab.getCell(row, 12).value = { formula: `J${row}-K${row}` }
    }
  })
  const lastAb = 13 + m.abstract.length
  const absTotal = lastAb + 1
  for (const row of [absTotal, absTotal + 1]) {
    for (let c = 1; c <= 12; c++)
      style(ab.getCell(row, c), { bold: true, fill: FILL_TOTAL, al: c === 1 ? CENTER : RIGHT, nf: '#,##0.00' })
    style(ab.getCell(row, 2), { bold: true, fill: FILL_TOTAL, al: MID_WRAP })
  }
  ab.getCell(absTotal, 2).value = 'Total'
  if (m.abstract.length) {
    ab.getCell(absTotal, 10).value = { formula: `SUM(J14:J${lastAb})` }
    ab.getCell(absTotal, 11).value = { formula: `SUM(K14:K${lastAb})` }
    ab.getCell(absTotal, 12).value = { formula: `J${absTotal}-K${absTotal}` }
  }
  ab.getCell(absTotal + 1, 2).value = 'Total as written in book'

  // ---------------- Seigniorage charges ----------------
  const sg = wb.addWorksheet('Seigniorage charges')
  setup(sg, 10, [8, 38, 10, 60, 34, 14, 8, 10, 10, 15], 0.75)
  title(sg, `MEASUREMENT BOOK ${mbName} | SEIGNIORAGE CHARGES`, 10, 14, 28)
  headRow(sg, 3, [
    'Chart Sl.',
    'Chart item',
    'Abstract item',
    'Abstract item description',
    'Sub item',
    'Quantity (from Abstract)',
    'Unit',
    'Factor',
    'Rate (Rs)',
    'Amount (Rs)'
  ])
  sg.getRow(3).height = 34
  let sr = 4
  // Listed in chart order (like the reference), not abstract order; items with no chart match are left out.
  m.abstract
    .map((a, i) => ({ a, i, chart: SEIG_CHART.find((c) => c.sl === a.chartSl) }))
    .filter((x) => x.chart)
    .sort((x, y) => x.chart!.sl - y.chart!.sl || x.i - y.i)
    .forEach(({ a, i, chart: c0 }) => {
      const chart = c0!
      const start = sr
      chart.subs.forEach((sub) => {
        for (let c = 1; c <= 10; c++) style(sg.getCell(sr, c), { al: { vertical: 'top' } })
        style(sg.getCell(sr, 1), { al: { horizontal: 'center', vertical: 'top' } })
        style(sg.getCell(sr, 2), { al: TOPWRAP })
        style(sg.getCell(sr, 3), { al: { horizontal: 'center', vertical: 'top' } })
        style(sg.getCell(sr, 4), { al: TOPWRAP })
        style(sg.getCell(sr, 5), { al: TOPWRAP })
        style(sg.getCell(sr, 6), { al: { vertical: 'top' }, nf: '0.00' })
        style(sg.getCell(sr, 7), { al: { horizontal: 'center', vertical: 'top' } })
        style(sg.getCell(sr, 8), { color: BLUE, al: { vertical: 'top' }, nf: '0.0000' })
        style(sg.getCell(sr, 9), { color: BLUE, al: { vertical: 'top' }, nf: '0.00' })
        style(sg.getCell(sr, 10), { al: { vertical: 'top' }, nf: '#,##0.00' })
        sg.getCell(sr, 5).value = sub.label ?? sub.material
        sg.getCell(sr, 6).value = { formula: `Abstract!D${14 + i}` }
        sg.getCell(sr, 7).value = { formula: `Abstract!G${14 + i}` }
        sg.getCell(sr, 8).value = sub.factor
        sg.getCell(sr, 9).value = SEIG_RATES[sub.material]
        sg.getCell(sr, 10).value = { formula: `ROUND(F${sr}*IF(H${sr}="",0,H${sr})*I${sr},2)` }
        sr++
      })
      // Chart Sl., chart item, abstract item and description span the item's sub rows (merged, like the reference).
      for (const c of [1, 2, 3, 4]) if (sr - 1 > start) sg.mergeCells(start, c, sr - 1, c)
      sg.getCell(start, 1).value = chart.sl
      sg.getCell(start, 2).value = chart.name
      sg.getCell(start, 3).value = String(i + 1)
      sg.getCell(start, 4).value = a.description
      sg.getRow(start).height = 75
    })
  const lastSg = sr - 1
  if (lastSg >= 4) {
    const total = sr
    for (let c = 1; c <= 10; c++) style(sg.getCell(total, c), { bold: true, fill: FILL_TOTAL, nf: '#,##0.00' })
    sg.getCell(total, 5).value = 'Total seigniorage charges'
    sg.getCell(total, 10).value = { formula: `SUM(J4:J${lastSg})` }
    const line = (row: number, label: string, share: number | null, formula: string) => {
      sg.mergeCells(row, 1, row, 5)
      for (let c = 1; c <= 10; c++) style(sg.getCell(row, c), { bold: true, color: GREEN, al: RIGHT, nf: '#,##0.00' })
      sg.getCell(row, 1).value = label
      sg.getCell(row, 1).alignment = { horizontal: 'right', vertical: 'middle', wrapText: true }
      if (share != null) {
        style(sg.getCell(row, 8), { color: BLUE, al: RIGHT, nf: '0.00' })
        sg.getCell(row, 8).value = share
      }
      sg.getCell(row, 10).value = { formula }
    }
    line(
      total + 1,
      'Seigniorage charges excluding Sand (Metal + M.sand only) calculating permit fee',
      null,
      `J${total}-SUMIF(E4:E${lastSg},"Sand*",J4:J${lastSg})`
    )
    line(
      total + 2,
      `Permit fee (${SEIG_PERMIT_SHARE * 100}% seigniorage charges excluding Sand)`,
      SEIG_PERMIT_SHARE,
      `ROUND(J${total + 1}*H${total + 2},2)`
    )
    line(
      total + 3,
      `SMET (${SEIG_SMET_SHARE * 100}% seigniorage charges)`,
      SEIG_SMET_SHARE,
      `ROUND(J${total}*H${total + 3},2)`
    )
    line(
      total + 4,
      `DMFT (${SEIG_DMFT_SHARE * 100}% seigniorage charges)`,
      SEIG_DMFT_SHARE,
      `ROUND(J${total}*H${total + 4},2)`
    )
    sg.getRow(total + 1).height = 34
  }

  // Text cells and the Sl/page/description columns keep General format, like the reference.
  wb.eachSheet((ws) =>
    ws.eachRow((row) =>
      row.eachCell((c) => {
        if (typeof c.value === 'string' || (ws.name === 'Measurements' && c.fullAddress.col <= 3)) c.numFmt = 'General'
      })
    )
  )

  return Buffer.from(await wb.xlsx.writeBuffer())
}

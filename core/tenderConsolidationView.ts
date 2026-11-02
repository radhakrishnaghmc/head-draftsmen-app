import { joinLines } from './tenderAgents/shared'

// Reads the portal's "Tender Consolidation View" page (the full tender notice:
// particulars, dates, "Required Tender Documents Details"…). The Evaluation Sheet
// takes two things from it: the tender's conditions (the documents required from
// the bidder) and the period of completion.

export interface TenderConsolidation {
  tenderId: string | null
  /** Period of completion in months (days are taken as 30 to a month). */
  completionMonths: number | null
  /** Documents required from the bidder, in order, with the portal's "1. " numbering stripped. */
  conditions: string[]
}

/** The period of completion in years (the "N" of 3AN−B): months / 12. */
export function completionYears(months: number | null | undefined): number | null {
  return months != null && months > 0 ? months / 12 : null
}

function readCompletionMonths(text: string): number | null {
  const at = text.search(/Period\s*of\s*Completion/i)
  if (at < 0) return null
  const near = text.slice(at, at + 160)
  const unit = /\b(Months?|Days?|Weeks?|Years?)\b/i.exec(near)?.[1].toLowerCase() ?? 'months'
  // The value sits between "(In" and the unit, or after the unit — skip digits inside the label itself.
  const after = near.replace(/^Period\s*of\s*Completion\s*\/?\s*Delivery\s*Period/i, '')
  const n = /(\d+(?:\.\d+)?)/.exec(after)?.[1]
  if (!n) return null
  const v = Number(n)
  if (!(v > 0)) return null
  if (unit.startsWith('day')) return v / 30
  if (unit.startsWith('week')) return (v * 7) / 30
  if (unit.startsWith('year')) return v * 12
  return v
}

function readConditions(text: string): string[] {
  const start = text.search(/Required\s*Tender\s*Documents\s*Details/i)
  if (start < 0) return []
  let body = text.slice(start).replace(/^Required\s*Tender\s*Documents\s*Details/i, '')
  const end = body.search(/Geographical\s*Particulars/i)
  if (end >= 0) body = body.slice(0, end)
  body = body.replace(/^\s*S\.?\s*No\s+Documents\s+Required\s+from\s+Contractor\s+Stage\s+Optional/i, '')
  // Every row ends "<STAGE> Mandatory|Optional"; its S.No may sit before or inside the wrapped text.
  const rows = body.split(/\b[A-Z]{3,}\s+(?:Mandatory|Optional)/)
  const out: string[] = []
  for (const row of rows) {
    const t = row
      .replace(/^\s*\d+\s+/, '') // S.No column
      .replace(/^\s*\d+\s*[.)]\s*/, '') // the portal's "2." numbering
      .replace(/\s+/g, ' ')
      .trim()
    if (t) out.push(t)
  }
  return out
}

/** Reads a Tender Consolidation View from its text lines (pdf.js or OCR). */
export function parseTenderConsolidationView(lines: string[]): TenderConsolidation {
  const text = joinLines(lines)
  const tenderId = /Tender\s*ID\s*[:\-]?\s*(\d{5,8})\b/i.exec(text)?.[1] ?? null
  return { tenderId, completionMonths: readCompletionMonths(text), conditions: readConditions(text) }
}

export interface PositionedItem {
  text: string
  x: number
  y: number
}

/**
 * The conditions read by column. The page prints them as a table (S.No | Document
 * | Stage | Optional) whose cells wrap and are vertically centred, so reading the
 * page line by line runs the S.No and "COMMON Mandatory" cells into the document
 * text. Here only the Document column is kept, and a new condition starts where
 * the line spacing opens up (rows are padded apart; wrapped lines are not).
 */
export function parseConditionsByColumn(items: PositionedItem[]): string[] {
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x)
  // Items may be whole phrases or single words, so find the headings by their first word.
  const start = sorted.findIndex((i, k) => /^Required/i.test(i.text) && /Tender/i.test(`${i.text} ${sorted[k + 1]?.text ?? ''}`))
  if (start < 0) return []
  const stop = sorted.findIndex((i, k) => k > start && /^Geographical/i.test(i.text))
  const region = sorted.slice(start + 1, stop > start ? stop : undefined)
  const docHead = region.find((i, k) => /^Documents/i.test(i.text) && /Required/i.test(`${i.text} ${region[k + 1]?.text ?? ''}`))
  const stageHead = region.find((i) => /^Stage$/i.test(i.text))
  if (!docHead) return []
  const left = docHead.x - 4
  const right = stageHead ? stageHead.x - 4 : docHead.x + 230
  const col = region.filter((i) => i.x >= left && i.x < right && i.y > docHead.y + 1)
  // cluster items sharing a baseline into lines, left to right
  const clusters: { y: number; items: PositionedItem[] }[] = []
  for (const it of col) {
    const last = clusters[clusters.length - 1]
    if (last && Math.abs(last.y - it.y) < 2) last.items.push(it)
    else clusters.push({ y: it.y, items: [it] })
  }
  const lines = clusters.map((c) => ({ y: c.y, text: c.items.sort((a, b) => a.x - b.x).map((i) => i.text).join(' ') }))
  const rows: string[][] = []
  lines.forEach((ln, k) => {
    const gap = k > 0 ? ln.y - lines[k - 1].y : Infinity
    if (gap > 11) rows.push([ln.text])
    else rows[rows.length - 1].push(ln.text)
  })
  return rows
    .map((r) => r.join(' ').replace(/^\s*\d+\s*[.)]\s*/, '').replace(/\s+/g, ' ').trim())
    .filter(Boolean)
}

/** A measurement book PDF read page by page into this model, then exported to Excel (core/measurementBookExcel.ts). */
export interface MbLine {
  location: string
  no: number | null
  l: number | null
  b: (number | null)[] // up to 3 breadth values
  d: (number | null)[] // up to 3 depth values
  /** Contents/quantity as written in the book for this line. */
  written: number | null
}

export interface MbItem {
  /** PDF page the item's heading is on. */
  mbPage: number
  description: string
  unit: string
  lines: MbLine[]
  totalWritten: number | null
  /** Where the item carries on to a later PDF page: that page, and how many lines precede it (shown as a "(contd.)" row). */
  contd: { page: number; atLine: number }[]
}

export interface MbAbstractRow {
  description: string
  qtyWritten: number | null
  unit: string
  rate: number | null
  amountWritten: number | null
  /** Seigniorage chart item (core/seigniorageChart.ts) this abstract item falls under, or null if none — those are left out. */
  chartSl: number | null
}

export interface MbHeader {
  work: string
  agency: string
  estimate: string
  contract: string
  percentage: string
  markOut: string
  completion: string
}

export interface MbModel {
  header: MbHeader
  items: MbItem[]
  abstract: MbAbstractRow[]
}

export const emptyMbModel = (): MbModel => ({
  header: { work: '', agency: '', estimate: '', contract: '', percentage: '', markOut: '', completion: '' },
  items: [],
  abstract: []
})

const num = (v: unknown): number | null => {
  const n = typeof v === 'string' ? Number(v.replace(/,/g, '').trim()) : v
  return typeof n === 'number' && Number.isFinite(n) && v !== '' ? n : null
}
const str = (v: unknown): string => (typeof v === 'string' ? v.trim() : v == null ? '' : String(v))
const three = (v: unknown): (number | null)[] => {
  const a = Array.isArray(v) ? v : [v]
  return [0, 1, 2].map((i) => num(a[i]))
}

/** Merge one page's JSON reply (from the per-page Claude prompt) into the model; returns false if it wasn't valid JSON. */
export function mergeMbPage(model: MbModel, page: number, reply: string): MbModel | null {
  const a = reply.indexOf('{')
  const z = reply.lastIndexOf('}')
  let r: any
  try {
    r = JSON.parse(reply.slice(a, z + 1))
  } catch {
    return null
  }
  const header = { ...model.header }
  for (const k of Object.keys(header) as (keyof MbHeader)[]) {
    const v = str(r?.header?.[k])
    if (v && !header[k]) header[k] = v
  }
  const items = model.items.map((i) => ({ ...i, lines: [...i.lines], contd: [...i.contd] }))
  for (const it of Array.isArray(r?.items) ? r.items : []) {
    const lines: MbLine[] = (Array.isArray(it?.lines) ? it.lines : []).map((l: any) => ({
      location: str(l?.location),
      no: num(l?.no),
      l: num(l?.l),
      b: three(l?.b),
      d: three(l?.d),
      written: num(l?.written)
    }))
    const last = items[items.length - 1]
    if (it?.continued && last) {
      last.contd.push({ page, atLine: last.lines.length })
      last.lines.push(...lines)
      last.totalWritten = num(it.totalWritten) ?? last.totalWritten
    } else {
      items.push({ mbPage: page, description: str(it?.description), unit: str(it?.unit), lines, totalWritten: num(it?.totalWritten), contd: [] })
    }
  }
  const abstract = [
    ...model.abstract,
    ...(Array.isArray(r?.abstract) ? r.abstract : []).map((x: any) => ({
      description: str(x?.description),
      qtyWritten: num(x?.qtyWritten),
      unit: str(x?.unit),
      rate: num(x?.rate),
      amountWritten: num(x?.amountWritten),
      chartSl: num(x?.chartSl)
    }))
  ]
  return { header, items, abstract }
}

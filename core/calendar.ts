// Parser for the Telangana Government 2026 holiday calendar page.
// The page renders 12 month tables; each row is
//   <th class="fix|opt">MMM DD</th><td>Holiday name</td>
// where class "fix" = general/public holiday and "opt" = optional holiday.

export type HolidayType = 'public' | 'optional'

export interface CalendarHoliday {
  /** Raw date label as shown, e.g. "Jan 26". */
  date: string
  name: string
  type: HolidayType
}

export interface CalendarMonth {
  month: string
  holidays: CalendarHoliday[]
}

export interface CalendarData {
  year: string
  months: CalendarMonth[]
  /** ISO timestamp of when the data was scraped. */
  fetchedAt: string
}

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December'
]

const MONTH_ABBR: Record<string, string> = {
  jan: 'January',
  feb: 'February',
  mar: 'March',
  apr: 'April',
  may: 'May',
  jun: 'June',
  jul: 'July',
  aug: 'August',
  sep: 'September',
  oct: 'October',
  nov: 'November',
  dec: 'December'
}

/** Ordered full month names, index 0 = January. */
export const MONTH_NAMES = MONTHS

/** Extract the day-of-month number from a label like "Jul 12" or "Jan 26 – 28". */
export function holidayDay(dateLabel: string): number | null {
  const m = dateLabel.match(/(\d{1,2})/)
  return m ? parseInt(m[1], 10) : null
}

export interface HolidayInfo {
  name: string
  type: HolidayType
}

/**
 * Build a day-of-month → holiday lookup for a given calendar month index
 * (0 = January). Matches the parsed month by its name.
 */
export function holidaysByDay(
  data: CalendarData,
  monthIndex: number
): Map<number, HolidayInfo> {
  const map = new Map<number, HolidayInfo>()
  const name = MONTHS[monthIndex]
  const entry = data.months.find((m) => m.month === name)
  if (!entry) return map
  for (const h of entry.holidays) {
    const day = holidayDay(h.date)
    if (day == null) continue
    // Prefer public over optional when both fall on the same day.
    const existing = map.get(day)
    if (!existing || (existing.type !== 'public' && h.type === 'public')) {
      map.set(day, { name: h.name, type: h.type })
    }
  }
  return map
}

function decodeEntities(s: string): string {
  return s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, "'")
    .replace(/&#8211;/g, '–')
    .replace(/&nbsp;/g, ' ')
    .replace(/&ndash;/g, '–')
}

function clean(s: string): string {
  return decodeEntities(s.replace(/<[^>]+>/g, ' '))
    .replace(/\s+/g, ' ')
    .trim()
}

/** Parse the calendar HTML into month-wise holiday lists. */
export function parseCalendarHtml(htmlText: string, year = '2026'): CalendarData {
  const tables = htmlText.match(/<table[\s\S]*?<\/table>/gi) ?? []
  const months: CalendarMonth[] = []

  const rowRe =
    /<th\b[^>]*class="(opt|fix)"[^>]*>([\s\S]*?)<\/th>\s*<td\b[^>]*>([\s\S]*?)<\/td>/gi

  for (let i = 0; i < tables.length; i++) {
    // Drop commented-out rows so we don't pick up stale/duplicate entries.
    const tableHtml = tables[i].replace(/<!--[\s\S]*?-->/g, '')
    const holidays: CalendarHoliday[] = []
    let m: RegExpExecArray | null
    rowRe.lastIndex = 0
    while ((m = rowRe.exec(tableHtml)) !== null) {
      const date = clean(m[2])
      const name = clean(m[3])
      if (!date || !name) continue
      holidays.push({
        date,
        name,
        type: m[1].toLowerCase() === 'fix' ? 'public' : 'optional'
      })
    }
    if (holidays.length === 0) continue

    // Month name: from the first holiday's abbrev, else fall back to order.
    const abbr = holidays[0].date.slice(0, 3).toLowerCase()
    const month = MONTH_ABBR[abbr] ?? MONTHS[months.length] ?? `Month ${i + 1}`
    months.push({ month, holidays })
  }

  return { year, months, fetchedAt: new Date().toISOString() }
}

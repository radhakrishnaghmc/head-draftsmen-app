import { useCallback, useEffect, useMemo, useState } from 'react'
import { api } from '../ipc'
import { IconTable, IconChevronLeft, IconChevronRight } from './Icons'
import type { CalendarData, HolidayType } from '@core/calendar'
import { MONTH_NAMES, holidaysByDay } from '@core/calendar'

const WEEKDAYS = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa']

/** The official Telangana government calendar this data is scraped from. */
const CALENDAR_SOURCE_URL = 'https://www.telangana.gov.in/downloads/calendar-2026/'

interface VisMonth {
  year: number
  mi: number
}

/** Absolute month ordinal so comparisons work across a year boundary. */
function ord(year: number, mi: number): number {
  return year * 12 + mi
}

function daysInMonth(year: number, mi: number): number {
  return new Date(year, mi + 1, 0).getDate()
}

function isSunday(year: number, mi: number, day: number): boolean {
  return new Date(year, mi, day).getDay() === 0
}

function isSecondSaturday(year: number, mi: number, day: number): boolean {
  return new Date(year, mi, day).getDay() === 6 && day >= 8 && day <= 14
}

/** Choose how many month grids to show based on the viewport width. */
function colsForWidth(w: number): number {
  if (w >= 1280) return 3
  if (w >= 880) return 2
  return 1
}

interface Props {
  cached: CalendarData | null
  onData: (d: CalendarData) => void
}

export default function Dashboard({ cached, onData }: Props) {
  const [data, setData] = useState<CalendarData | null>(cached)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  // The clicked "day 1" anchor for the working-day counter.
  const [anchor, setAnchor] = useState<{ year: number; mi: number; day: number } | null>(
    null
  )
  // Tender window: count up to the 3rd or 7th day. The final day (3 or 7) may
  // not land on a public holiday — it rolls to the next non-holiday day.
  const [tenderDays, setTenderDays] = useState<3 | 7>(7)

  const load = useCallback(
    async (force = false) => {
      setLoading(true)
      setError(null)
      try {
        const d = await api.fetchCalendar(force)
        setData(d)
        onData(d)
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e))
      } finally {
        setLoading(false)
      }
    },
    [onData]
  )

  useEffect(() => {
    // Use the in-memory copy if we already have it; otherwise read the
    // on-disk cache (no network). Refresh forces a fresh fetch.
    if (!data) load(false)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // How many months fit, based on window width (1 / 2 / 3).
  const [monthCols, setMonthCols] = useState(() => colsForWidth(window.innerWidth))
  useEffect(() => {
    const onResize = () => setMonthCols(colsForWidth(window.innerWidth))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  // Navigation offset (in months) from the current month.
  const [monthOffset, setMonthOffset] = useState(0)

  // The months to display, starting from the current month + offset.
  const visMonths = useMemo<VisMonth[]>(() => {
    const now = new Date()
    return Array.from({ length: monthCols }, (_, i) => {
      const d = new Date(now.getFullYear(), now.getMonth() + monthOffset + i, 1)
      return { year: d.getFullYear(), mi: d.getMonth() }
    })
  }, [monthCols, monthOffset])

  // Public-holiday day maps per visible month (used to skip while numbering).
  const holidayMaps = useMemo(() => {
    const maps = new Map<string, ReturnType<typeof holidaysByDay>>()
    if (!data) return maps
    for (const m of visMonths) {
      maps.set(`${m.year}-${m.mi}`, holidaysByDay(data, m.mi))
    }
    return maps
  }, [data, visMonths])

  // Working-day numbers: the anchor is day 0, following days count up. Every day
  // is counted — INCLUDING public holidays, Sundays and second Saturdays —
  // except that the final day of the window (the 3rd or 7th, per the selected
  // tender type) may not land on any of those non-working days: if it would,
  // that day is skipped so the number rolls forward to the next working day.
  // Numbered 0 through `tenderDays`.
  const numbers = useMemo(() => {
    const result = new Map<string, number>()
    if (!anchor) return result
    const anchorOrd = ord(anchor.year, anchor.mi)
    let count = 0
    for (const m of visMonths) {
      const hol = holidayMaps.get(`${m.year}-${m.mi}`)
      const total = daysInMonth(m.year, m.mi)
      const mOrd = ord(m.year, m.mi)
      for (let d = 1; d <= total; d++) {
        const before = mOrd < anchorOrd || (mOrd === anchorOrd && d < anchor.day)
        if (before) continue
        if (count > tenderDays) return result
        const isNonWorking =
          hol?.get(d)?.type === 'public' || isSunday(m.year, m.mi, d) || isSecondSaturday(m.year, m.mi, d)
        // The final day cannot be a non-working day — skip so the count rolls
        // forward to the next working day. Earlier days count regardless.
        if (count === tenderDays && isNonWorking) continue
        result.set(`${m.year}-${m.mi}-${d}`, count)
        count++
      }
    }
    return result
  }, [anchor, visMonths, holidayMaps, tenderDays])

  function toggleAnchor(year: number, mi: number, day: number, isPublic: boolean) {
    if (isPublic) return // can't start a working-day count on a holiday
    setAnchor((prev) =>
      prev && prev.year === year && prev.mi === mi && prev.day === day
        ? null
        : { year, mi, day }
    )
  }

  return (
    <div className="dash">
      {error && !data && (
        <div className="card">
          <div className="empty">
            <IconTable />
            <p>{error}</p>
            <button className="primary" onClick={() => load(true)}>
              Try again
            </button>
          </div>
        </div>
      )}

      {error && data && <div className="notice warn">Showing saved copy — {error}</div>}

      {data && (
        <>
          {/* Three-month working-day counter */}
          <section className="card cal-counter">
            <div className="cal-counter-head">
              <div>
                <div className="cal-tender-modes">
                  <label className="cal-tender-mode">
                    <input
                      type="checkbox"
                      checked={tenderDays === 3}
                      onChange={() => setTenderDays(3)}
                    />
                    3-day tender
                  </label>
                  <label className="cal-tender-mode">
                    <input
                      type="checkbox"
                      checked={tenderDays === 7}
                      onChange={() => setTenderDays(7)}
                    />
                    7-day tender
                  </label>
                </div>
              </div>
              <div className="cal-nav">
                {anchor && (
                  <button className="ghost" onClick={() => setAnchor(null)}>
                    Clear
                  </button>
                )}
                <button
                  className="cal-nav-btn"
                  title="Previous month"
                  onClick={() => setMonthOffset((o) => o - 1)}
                >
                  <IconChevronLeft />
                </button>
                {monthOffset !== 0 && (
                  <button className="cal-nav-today" onClick={() => setMonthOffset(0)}>
                    Today
                  </button>
                )}
                <button
                  className="cal-nav-btn"
                  title="Next month"
                  onClick={() => setMonthOffset((o) => o + 1)}
                >
                  <IconChevronRight />
                </button>
              </div>
            </div>
            <div
              className="cal-months"
              style={{
                gridTemplateColumns: `repeat(${visMonths.length}, minmax(0, 1fr))`
              }}
            >
              {visMonths.map((m) => {
                const hol = holidayMaps.get(`${m.year}-${m.mi}`)
                const total = daysInMonth(m.year, m.mi)
                const lead = new Date(m.year, m.mi, 1).getDay()
                const cells: (number | null)[] = [
                  ...Array(lead).fill(null),
                  ...Array.from({ length: total }, (_, i) => i + 1)
                ]
                return (
                  <div className="cal-month" key={`${m.year}-${m.mi}`}>
                    <div className="cal-month-title">
                      {MONTH_NAMES[m.mi]} {m.year}
                    </div>
                    <div className="cal-weekdays">
                      {WEEKDAYS.map((w) => (
                        <span key={w}>{w}</span>
                      ))}
                    </div>
                    <div className="cal-days">
                      {cells.map((day, idx) => {
                        if (day == null)
                          return <span key={idx} className="cal-day empty" />
                        const info = hol?.get(day)
                        const type: HolidayType | undefined = info?.type
                        const num = numbers.get(`${m.year}-${m.mi}-${day}`)
                        const isAnchor =
                          anchor &&
                          anchor.year === m.year &&
                          anchor.mi === m.mi &&
                          anchor.day === day
                        const cls = [
                          'cal-day',
                          type ? `hol-${type}` : '',
                          !type && (isSunday(m.year, m.mi, day) || isSecondSaturday(m.year, m.mi, day))
                            ? 'sunday'
                            : '',
                          num != null ? 'numbered' : '',
                          isAnchor ? 'anchor' : ''
                        ]
                          .filter(Boolean)
                          .join(' ')
                        return (
                          <button
                            key={idx}
                            className={cls}
                            title={info ? `${info.name} (${info.type})` : undefined}
                            onClick={() =>
                              toggleAnchor(m.year, m.mi, day, type === 'public')
                            }
                          >
                            <span className="cal-date">{day}</span>
                            {num != null && <span className="cal-num">{num}</span>}
                          </button>
                        )
                      })}
                    </div>
                  </div>
                )
              })}
            </div>
            <div className="cal-legend">
              <span className="lg lg-public">Public holiday</span>
              <span className="lg lg-optional">Optional holiday</span>
              <span className="lg lg-num">Working-day number</span>
              <button
                className="cal-source-link"
                onClick={() => api.openPath(CALENDAR_SOURCE_URL)}
                title="Open the official Telangana government calendar"
              >
                View original calendar ↗
              </button>
            </div>
          </section>
        </>
      )}
    </div>
  )
}

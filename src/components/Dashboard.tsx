import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { api } from '../ipc'
import { IconTable, IconChevronLeft, IconChevronRight } from './Icons'
import type { CalendarData, HolidayType } from '@core/calendar'
import { MONTH_NAMES, holidaysByDay } from '@core/calendar'
import { type Office, officeScopedKey, isZoneOnlyOffice, CONTACT_KEYS } from '../office'

const WEEKDAYS = ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa']

// The government publishes one calendar page per year (…/calendar-2026/,
// …/calendar-2027/ once that year's page exists) rather than one page that
// stays current — so instead of this app guessing/hardcoding a year, the
// source link is user-editable and remembered here. Update it once a new
// year's page goes up (usually announced around December).
const CALENDAR_URL_KEY = 'calendarSourceUrl'
const DEFAULT_CALENDAR_URL = 'https://www.telangana.gov.in/downloads/calendar-2026/'

/**
 * Guesses a future/past year's page from the last known-good one by swapping
 * its 4-digit year — the government's own URLs are exactly this pattern
 * (…/calendar-2026/, …/calendar-2027/), so this holds until they restructure
 * the site. When the guess is wrong (page not published yet, or the pattern
 * changed), the fetch simply fails and the caller shows the "change the
 * link" prompt instead of silently displaying nothing.
 */
function urlForYear(url: string, year: number): string {
  return /\d{4}/.test(url) ? url.replace(/\d{4}/, String(year)) : url
}

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
  /** The chosen office — determines whether the contact field below the tender-mode toggle reads "Executive Engineer Phone" or "Superintending Engineer Phone", and scopes where its value (and the Head Draughtsman phone) are remembered. */
  office?: Office
}

const CURRENT_YEAR = new Date().getFullYear()

export default function Dashboard({ cached, onData, office }: Props) {
  // Per-year, not one blob: the month grid can show two or three months at
  // once and, near a year boundary, those can span two different years —
  // each fetched (and failed/loading) independently so December's data
  // never gets blamed for January's missing page or vice versa.
  const [yearData, setYearData] = useState<Map<number, CalendarData>>(() => {
    const m = new Map<number, CalendarData>()
    if (cached) m.set(Number(cached.year), cached)
    return m
  })
  const [yearError, setYearError] = useState<Map<number, string>>(new Map())
  const [yearLoading, setYearLoading] = useState<Set<number>>(new Set())
  // Years already requested (loaded, loading, or failed) — a ref, not state,
  // purely to stop the visible-months effect below from re-requesting a year
  // it already asked for; it shouldn't itself trigger a re-render.
  const requestedYears = useRef<Set<number>>(new Set(yearData.keys()))
  const [sourceUrl, setSourceUrl] = useState(() => localStorage.getItem(CALENDAR_URL_KEY) || DEFAULT_CALENDAR_URL)
  const [editingUrl, setEditingUrl] = useState<number | null>(null)
  const [urlDraft, setUrlDraft] = useState('')
  // The clicked "day 1" anchor for the working-day counter.
  const [anchor, setAnchor] = useState<{ year: number; mi: number; day: number } | null>(
    null
  )
  // Tender window: count up to the 3rd or 7th day. The final day (3 or 7) may
  // not land on a public holiday — it rolls to the next non-holiday day.
  const [tenderDays, setTenderDays] = useState<3 | 7>(7)

  // The office's own Engineer phone (EE for a Circle office, SE for a Zone-only
  // office) and Head Draughtsman phone — shown here, always visible, so they
  // don't have to be re-typed inside the "Issue Tender Notice" dialog every
  // time. Shares the same office-scoped localStorage keys as that dialog (see
  // TenderNoticeButton's CONTACT_KEYS), so entering it in either place fills
  // the other. Reloaded whenever the chosen office changes.
  const [enginPhone, setEnginPhone] = useState('')
  const [hdPhone, setHdPhone] = useState('')
  useEffect(() => {
    setEnginPhone(localStorage.getItem(officeScopedKey(CONTACT_KEYS.eePhone, office)) || '')
    setHdPhone(localStorage.getItem(officeScopedKey(CONTACT_KEYS.hdPhone, office)) || '')
  }, [office])
  function saveEnginPhone(v: string) {
    setEnginPhone(v)
    localStorage.setItem(officeScopedKey(CONTACT_KEYS.eePhone, office), v.trim())
  }
  function saveHdPhone(v: string) {
    setHdPhone(v)
    localStorage.setItem(officeScopedKey(CONTACT_KEYS.hdPhone, office), v.trim())
  }

  // Fetches one specific year's calendar, guessing its URL from the last
  // known-good link (urlForYear) unless an explicit override is given (the
  // user pasting in a corrected link for a year that guessed wrong).
  const loadYear = useCallback(
    async (year: number, force = false, overrideUrl?: string) => {
      requestedYears.current.add(year)
      setYearLoading((prev) => new Set(prev).add(year))
      setYearError((prev) => {
        const next = new Map(prev)
        next.delete(year)
        return next
      })
      try {
        const url = overrideUrl ?? urlForYear(sourceUrl, year)
        const d = await api.fetchCalendar(url, force)
        setYearData((prev) => new Map(prev).set(Number(d.year), d))
        if (overrideUrl) {
          // A manually-fixed link becomes the new template other years guess
          // from — so once the government's actual 2027 page is pasted in,
          // guessing 2028 next year starts from a link that's known to work.
          localStorage.setItem(CALENDAR_URL_KEY, overrideUrl)
          setSourceUrl(overrideUrl)
        }
        if (year === CURRENT_YEAR) onData(d)
      } catch (e) {
        setYearError((prev) => new Map(prev).set(year, e instanceof Error ? e.message : String(e)))
      } finally {
        setYearLoading((prev) => {
          const next = new Set(prev)
          next.delete(year)
          return next
        })
      }
    },
    [onData, sourceUrl]
  )

  useEffect(() => {
    // Use the in-memory copy if we already have it; otherwise read the
    // on-disk cache (no network).
    if (!yearData.has(CURRENT_YEAR)) void loadYear(CURRENT_YEAR, false)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // A pasted link fixes exactly the one year it points at — Change Year's
  // Link is opened per-month (see editingUrl), so this never has to guess
  // which year the user meant.
  function saveSourceUrl(year: number, url: string) {
    const trimmed = url.trim()
    setEditingUrl(null)
    if (!trimmed) return
    void loadYear(year, true, trimmed)
  }

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

  // Whenever navigation brings a not-yet-seen year into view (paging across
  // a year boundary), guess and fetch its page automatically — this is what
  // makes the January-2027 page load without the user having to do anything,
  // as long as the guessed URL is right.
  useEffect(() => {
    for (const m of visMonths) {
      if (!requestedYears.current.has(m.year)) void loadYear(m.year, false)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [visMonths])

  // Public-holiday day maps per visible month (used to skip while numbering)
  // — each month's own year's data, never another year's mistaken for it.
  const holidayMaps = useMemo(() => {
    const maps = new Map<string, ReturnType<typeof holidaysByDay>>()
    for (const m of visMonths) {
      const d = yearData.get(m.year)
      if (d) maps.set(`${m.year}-${m.mi}`, holidaysByDay(d, m.mi))
    }
    return maps
  }, [yearData, visMonths])

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

  // Opens the per-year link editor, prefilled with a guessed URL for that
  // year so fixing it is usually just "paste the real one over this".
  function openLinkEditor(year: number) {
    setUrlDraft(urlForYear(sourceUrl, year))
    setEditingUrl(year)
  }

  const primaryError = yearError.get(CURRENT_YEAR)
  const hasAnyData = yearData.size > 0

  return (
    <div className="dash">
      {primaryError && !yearData.has(CURRENT_YEAR) && (
        <div className="card">
          <div className="empty">
            <IconTable />
            <p>Couldn't load {CURRENT_YEAR}'s calendar automatically — {primaryError}</p>
            <div className="boq-actions">
              <button className="primary" onClick={() => loadYear(CURRENT_YEAR, true)}>
                Try again
              </button>
              <button className="ghost" onClick={() => openLinkEditor(CURRENT_YEAR)}>
                Change year's link
              </button>
            </div>
          </div>
        </div>
      )}

      {[...yearError].filter(([y]) => yearData.has(y)).map(([y, msg]) => (
        <div className="notice warn" key={y}>
          Keeping the previous {y} calendar — {msg}
        </div>
      ))}

      {hasAnyData && (
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
                  <label className="cal-tender-phone">
                    <span>{isZoneOnlyOffice(office) ? 'SE Phone' : 'EE Phone'}</span>
                    <input
                      type="tel"
                      value={enginPhone}
                      onChange={(e) => saveEnginPhone(e.target.value)}
                      placeholder="10-digit mobile number"
                    />
                  </label>
                  <label className="cal-tender-phone">
                    <span>HD Phone</span>
                    <input
                      type="tel"
                      value={hdPhone}
                      onChange={(e) => saveHdPhone(e.target.value)}
                      placeholder="10-digit mobile number"
                    />
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
                // No calendar page has loaded for this month's year (still
                // fetching, the guessed URL was wrong, or nothing published
                // yet) — show that instead of a blank/misleading grid.
                if (!yearData.has(m.year)) {
                  const err = yearError.get(m.year)
                  const busy = yearLoading.has(m.year)
                  return (
                    <div className="cal-month" key={`${m.year}-${m.mi}`}>
                      <div className="cal-month-title">
                        {MONTH_NAMES[m.mi]} {m.year}
                      </div>
                      <div className="cal-month-missing">
                        {busy ? (
                          <p>Loading {m.year}'s calendar…</p>
                        ) : (
                          <>
                            <p>
                              {err
                                ? `Couldn't load ${m.year}'s calendar — ${err}`
                                : `No calendar loaded for ${m.year} yet.`}
                            </p>
                            <button className="ghost" onClick={() => openLinkEditor(m.year)}>
                              Change year's link
                            </button>
                          </>
                        )}
                      </div>
                    </div>
                  )
                }
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
              {editingUrl != null ? (
                <form
                  className="cal-source-edit"
                  onSubmit={(e) => {
                    e.preventDefault()
                    saveSourceUrl(editingUrl, urlDraft)
                  }}
                >
                  <span className="cal-source-edit-label">{editingUrl} link:</span>
                  <input
                    type="url"
                    autoFocus
                    value={urlDraft}
                    placeholder={`https://www.telangana.gov.in/downloads/calendar-${editingUrl}/`}
                    onChange={(e) => setUrlDraft(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Escape') setEditingUrl(null)
                    }}
                  />
                  <button type="submit" className="cal-source-save">
                    Save
                  </button>
                  <button type="button" className="cal-source-link" onClick={() => setEditingUrl(null)}>
                    Cancel
                  </button>
                </form>
              ) : (
                <>
                  <button
                    className="cal-source-link"
                    onClick={() => api.openPath(sourceUrl)}
                    title="Open the official Telangana government calendar"
                  >
                    View original calendar ↗
                  </button>
                  <button
                    className="cal-source-link"
                    onClick={() => openLinkEditor(CURRENT_YEAR)}
                    title="Paste a corrected calendar page link for a year that failed to auto-load"
                  >
                    Change year's link
                  </button>
                </>
              )}
            </div>
          </section>
        </>
      )}
    </div>
  )
}

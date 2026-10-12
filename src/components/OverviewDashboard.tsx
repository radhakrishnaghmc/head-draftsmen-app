import { useMemo, useState, type SVGProps } from 'react'
import type { TabKey } from './Sidebar'
import type { Office } from '../office'
import type { TodoItem, MBScrutinyItem, TenderReminder } from '@core/types'
import type { MonitoringFormatSummary, MfStatusKey } from '@core/monitoringFormat'
import { indianDigitGroups } from '@core/worksAmounts'
import WeatherCard from './WeatherCard'
import {
  IconChecklist,
  IconEye,
  IconSearch,
  IconPlus,
  IconDoc,
  IconCheck,
  IconBolt,
  IconCalendar,
  IconClipboard,
  IconWarn,
  IconClose,
  IconRefresh
} from './Icons'

interface Props {
  office: Office
  monitoringFormat?: MonitoringFormatSummary
  todos: TodoItem[]
  mbScrutiny: MBScrutinyItem[]
  tenderReminders: TenderReminder[]
  cementSteelHasNew: boolean
  onNavigate: (tab: TabKey) => void
  onExportStatusWorks: (key: MfStatusKey, label: string) => Promise<boolean>
  onExportWardWorks: (
    wardLabel: string,
    pivotKind: 'ward' | 'circle',
    statusKey?: MfStatusKey | null,
    statusLabel?: string | null
  ) => Promise<boolean>
}

// Fixed draw order — never reassigned per data, so a color always means the
// same stage. Completed / Held Up / Cancelled reuse the app's reserved status
// colors (good / warning / critical); the three in-between stages each get
// their own categorical hue from the app's existing brand palette.
const MF_TILES: {
  label: string
  key: MfStatusKey
  color: string
  icon: (p: SVGProps<SVGSVGElement>) => JSX.Element
}[] = [
  { label: 'Completed', key: 'completed', color: 'var(--ok)', icon: IconCheck },
  { label: 'In Progress', key: 'progressTotal', color: 'var(--teal-500)', icon: IconBolt },
  { label: 'To Be Started', key: 'toBeStarted', color: 'var(--violet-500)', icon: IconCalendar },
  { label: 'Tender Process', key: 'tenderProcess', color: 'var(--sky-500)', icon: IconClipboard },
  { label: 'Held Up', key: 'heldUp', color: 'var(--warn)', icon: IconWarn },
  { label: 'Cancelled', key: 'cancelled', color: 'var(--danger)', icon: IconClose }
]

const DONUT_R = 50
const DONUT_C = 2 * Math.PI * DONUT_R

// Wards/circles rank by work count, not identity (the set is whatever the
// office's own Monitoring Format workbook defines, never a fixed category
// list) — so this is a magnitude encoding, not a categorical one. A single
// hue light-to-dark (one indigo ramp, largest ward darkest) is the correct
// form for that job; a full hue rotation across wards would read as
// "rainbow" with no actual meaning behind the different colors.
function wardColor(rank: number, total: number): string {
  const t = total <= 1 ? 0 : rank / (total - 1)
  const lightness = Math.round(72 - t * 40)
  return `hsl(226, 62%, ${lightness}%)`
}

function formatDDMMYYYY(iso: string): string {
  if (!iso) return ''
  const [y, m, d] = iso.split('-')
  return `${d}.${m}.${y}`
}

function deadlineTab(kind: 'todo' | 'mb' | 'tender'): TabKey {
  if (kind === 'mb') return 'mbScrutiny'
  if (kind === 'tender') return 'search'
  return 'todo'
}

interface Deadline {
  id: string
  text: string
  dueLabel: string
  sortTs: number
  kind: 'todo' | 'mb' | 'tender'
}

/**
 * Home base for the app: a snapshot of the current office's own data (Works
 * List size, To Do / MB Scrutiny progress) rather than a generic template —
 * there's no team or time-tracking concept in a single-user desktop tool, so
 * those reference-design widgets are replaced with the app's real per-office
 * lists instead of being faked.
 */
export default function OverviewDashboard({
  office,
  monitoringFormat,
  todos,
  mbScrutiny,
  tenderReminders,
  cementSteelHasNew,
  onNavigate,
  onExportStatusWorks,
  onExportWardWorks
}: Props) {
  const officeLabel = office.circle || office.zone || office.corporation || ''

  // Per-tile Excel-download feedback: the click triggers a native save dialog
  // with no visible page change of its own, so the tile itself shows a spinner
  // while it's in flight and a brief checkmark once the file is written (a
  // cancelled dialog just falls back to idle, no error needed).
  const [exportingKey, setExportingKey] = useState<MfStatusKey | null>(null)
  const [exportedKey, setExportedKey] = useState<MfStatusKey | null>(null)

  async function handleExportTile(key: MfStatusKey, label: string) {
    if (exportingKey) return
    setExportingKey(key)
    setExportedKey(null)
    try {
      const saved = await onExportStatusWorks(key, label)
      if (saved) {
        setExportedKey(key)
        setTimeout(() => setExportedKey((k) => (k === key ? null : k)), 2000)
      }
    } finally {
      setExportingKey((k) => (k === key ? null : k))
    }
  }

  const [exportingWard, setExportingWard] = useState<string | null>(null)
  const [exportedWard, setExportedWard] = useState<string | null>(null)
  const [wardStatusFilter, setWardStatusFilter] = useState<MfStatusKey | ''>('')

  async function handleExportWardTile(wardLabel: string) {
    if (exportingWard) return
    setExportingWard(wardLabel)
    setExportedWard(null)
    try {
      const statusLabel = MF_TILES.find((t) => t.key === wardStatusFilter)?.label
      const saved = await onExportWardWorks(
        wardLabel,
        monitoringFormat?.wardPivotKind === 'circle' ? 'circle' : 'ward',
        wardStatusFilter || null,
        statusLabel
      )
      if (saved) {
        setExportedWard(wardLabel)
        setTimeout(() => setExportedWard((w) => (w === wardLabel ? null : w)), 2000)
      }
    } finally {
      setExportingWard((w) => (w === wardLabel ? null : w))
    }
  }

  const mfTotals = monitoringFormat?.totals

  const wardTypes = useMemo(() => {
    const rows = (monitoringFormat?.wardRows || []).filter((r) => r.totalWorks.no > 0)
    const sorted = [...rows].sort((a, b) => b.totalWorks.no - a.totalWorks.no)
    return sorted.map((row, i) => ({ row, color: wardColor(i, sorted.length) }))
  }, [monitoringFormat])

  // Binned tasks/MBs are hidden from the dashboard entirely — they're neither
  // pending nor completed work anymore, just recoverable clutter.
  const activeTodos = useMemo(() => todos.filter((t) => !t.deletedAt), [todos])
  const activeMb = useMemo(() => mbScrutiny.filter((m) => !m.deletedAt), [mbScrutiny])

  const completedCount = activeTodos.filter((t) => t.done).length + activeMb.filter((m) => m.done).length
  const pendingCount = activeTodos.filter((t) => !t.done).length + activeMb.filter((m) => !m.done).length

  const totalTracked = completedCount + pendingCount
  const progressPct = totalTracked === 0 ? 0 : Math.round((completedCount / totalTracked) * 100)


  const deadlines: Deadline[] = useMemo(() => {
    // Sorted by a best-effort timestamp — To Do/MB dates are ISO so this is
    // exact; tender bid-closing times come from the portal as free text, so an
    // unparseable one just sorts to the end rather than being dropped.
    const items: Deadline[] = [
      ...activeTodos
        .filter((t) => !t.done && t.targetDate)
        .map((t) => ({
          id: t.id,
          text: t.text,
          dueLabel: formatDDMMYYYY(t.targetDate),
          sortTs: Date.parse(t.targetDate),
          kind: 'todo' as const
        })),
      ...activeMb
        .filter((m) => !m.done && m.targetDate)
        .map((m) => ({
          id: m.id,
          text: `MB No. ${m.mbNo} — ${m.agencyName}`,
          dueLabel: formatDDMMYYYY(m.targetDate as string),
          sortTs: Date.parse(m.targetDate as string),
          kind: 'mb' as const
        })),
      // Every reminder shows up here, not just ones with a resolved bid-closing
      // date — a pending/not-found lookup still gets an entry (dueLabel
      // explains the state, sortTs pushes it after anything with a real date)
      // so it isn't silently invisible on the dashboard while still listed on
      // the Search Tender tab's own Tender Reminders card.
      ...tenderReminders.flatMap((r) =>
        r.items.length > 0
          ? r.items.map((it, i) => ({
              id: `${r.id}-${i}`,
              text: it.workName || r.nitNo,
              dueLabel: it.bidClosing ? `Bid closing ${it.bidClosing}` : 'Bid closing time pending',
              sortTs: it.bidClosing ? Date.parse(it.bidClosing) : Infinity,
              kind: 'tender' as const
            }))
          : [
              {
                id: r.id,
                text: r.nitNo,
                dueLabel: r.status === 'pending' ? 'Looking up bid closing time…' : 'Not found on the portal yet',
                sortTs: Infinity,
                kind: 'tender' as const
              }
            ]
      )
    ]
    return items
      .map((d) => (Number.isNaN(d.sortTs) ? { ...d, sortTs: Infinity } : d))
      .sort((a, b) => a.sortTs - b.sortTs)
  }, [activeTodos, activeMb, tenderReminders])

  const nextDeadline = deadlines[0]
  const upcoming = deadlines.slice(0, 5)

  // Is the nearest deadline already past? Infinity (undated tender lookups)
  // never counts as overdue.
  const nextOverdue =
    !!nextDeadline && Number.isFinite(nextDeadline.sortTs) && nextDeadline.sortTs < Date.now()

  const recentMb = useMemo(
    () => [...activeMb].sort((a, b) => b.receivedDate.localeCompare(a.receivedDate)).slice(0, 4),
    [activeMb]
  )

  return (
    <div className="overview-dashboard">
      <div className="ov-stat-row">
        <div className="ov-stat ov-stat-hero">
          <span className="ov-stat-top">Total works</span>
          <b>{mfTotals ? mfTotals.totalWorks.no : '—'}</b>
          <span className="ov-stat-sub">
            {mfTotals ? monitoringFormat!.officeLabel : 'Import a Monitoring Format to populate'}
          </span>
        </div>
        <div className="ov-stat">
          <span className="ov-stat-top">Pending items</span>
          <b>{pendingCount}</b>
          <span className="ov-stat-sub">To Do &amp; MB Scrutiny open</span>
        </div>
        <div className="ov-stat">
          <span className="ov-stat-top">Completed</span>
          <b>{progressPct}%</b>
          <span className="ov-stat-sub">{completedCount} of {totalTracked} done</span>
        </div>
        <div className={`ov-stat ov-stat-nextdeadline${nextOverdue ? ' ov-stat-warn' : ''}`}>
          <span className="ov-stat-top">Next deadline</span>
          <b className="ov-stat-deadline">{nextDeadline ? nextDeadline.dueLabel : '—'}</b>
          <span className="ov-stat-sub">{nextDeadline ? nextDeadline.text : 'Nothing scheduled'}</span>
        </div>
      </div>
      <div className="ov-mf-row">
      <div className="card ov-card ov-mf-status">
        <div className="card-head">
          <span className="ov-mf-titles">
            <h3>Monitoring Format — Works Status</h3>
            {mfTotals && (
              <span className="sub">
                {monitoringFormat!.officeLabel}
                {monitoringFormat!.asOfDate ? ` · As of ${monitoringFormat!.asOfDate}` : ''}
              </span>
            )}
          </span>
          <button className="ghost ov-new-btn" onClick={() => onNavigate('data')}>
            {mfTotals ? 'Update' : 'Import'}
          </button>
        </div>
        {!mfTotals ? (
          <div className="empty">
            <p>Import the office's Monitoring Format workbook (Works List page) to see works status here.</p>
          </div>
        ) : (
          <div className="ov-mf-body">
            <div className="ov-mf-donut-col">
              <div className="ov-mf-donut-wrap">
                <svg viewBox="0 0 120 120" className="ov-donut ov-mf-donut">
                  <circle cx="60" cy="60" r={DONUT_R} className="ov-donut-track" />
                  {(() => {
                    let cursor = 0
                    return MF_TILES.map(({ key, color }) => {
                      const no = mfTotals[key].no
                      const share = mfTotals.totalWorks.no > 0 ? no / mfTotals.totalWorks.no : 0
                      if (share <= 0) return null
                      const segLen = share * DONUT_C
                      const gap = 2
                      const dash = `${Math.max(segLen - gap, 0)} ${DONUT_C - segLen + gap}`
                      const offset = -cursor
                      cursor += segLen
                      return (
                        <circle
                          key={key}
                          cx="60"
                          cy="60"
                          r={DONUT_R}
                          className="ov-mf-donut-seg"
                          style={{ stroke: color, strokeDasharray: dash, strokeDashoffset: offset }}
                        />
                      )
                    })
                  })()}
                </svg>
                <div className="ov-donut-text">
                  <b>{mfTotals.totalWorks.no}</b>
                  <span>Total Works</span>
                </div>
              </div>

              <ul className="ov-donut-legend-list">
                {MF_TILES.filter(({ key }) => mfTotals[key].no > 0).map(({ label, key, color }) => (
                  <li key={key}>
                    <i className="ov-donut-legend-dot" style={{ background: color }} />
                    <span>{label}</span>
                    <b>{mfTotals[key].no}</b>
                  </li>
                ))}
              </ul>
            </div>

            <div className="ov-mf-tiles">
              {MF_TILES.map(({ label, key, color, icon: Icon }) => {
                const bucket = mfTotals[key]
                const pct = mfTotals.totalWorks.no > 0 ? Math.round((bucket.no / mfTotals.totalWorks.no) * 100) : 0
                const isExporting = exportingKey === key
                const isExported = exportedKey === key
                return (
                  <button
                    key={key}
                    type="button"
                    className={`ov-mf-tile${isExporting ? ' ov-mf-tile-busy' : ''}${isExported ? ' ov-mf-tile-done' : ''}`}
                    style={{ ['--tile-color' as string]: color, ['--tile-pct' as string]: `${pct}%` }}
                    disabled={bucket.no === 0 || isExporting}
                    title={bucket.no > 0 ? `Download Excel of ${label} works` : undefined}
                    onClick={() => handleExportTile(key, label)}
                  >
                    <span className="ov-mf-tile-ic">
                      {isExporting ? <IconRefresh className="ov-mf-tile-spin" /> : isExported ? <IconCheck /> : <Icon />}
                    </span>
                    <span className="ov-mf-tile-main">
                      <span className="ov-mf-tile-label">
                        {label}
                        {isExporting && <span className="ov-mf-tile-status"> · Downloading…</span>}
                        {isExported && <span className="ov-mf-tile-status ov-mf-tile-status-done"> · Downloaded</span>}
                      </span>
                      <span className="ov-mf-tile-nums">
                        <b className="ov-mf-tile-no">{bucket.no}</b>
                        <span className="ov-mf-tile-pct">{pct}%</span>
                      </span>
                      <span className="ov-mf-tile-amt">Rs. {indianDigitGroups(bucket.amt)} L</span>
                    </span>
                  </button>
                )
              })}
            </div>
          </div>
        )}
      </div>

        <div className="card ov-card ov-ward-status">
          <div className="card-head">
            <span className="ov-mf-titles">
              <h3>{monitoringFormat?.wardPivotKind === 'circle' ? 'Circle Wise' : 'Ward Wise'}</h3>
              {mfTotals && (
                <span className="sub">
                  {monitoringFormat!.officeLabel} ·{' '}
                  {wardStatusFilter ? MF_TILES.find((t) => t.key === wardStatusFilter)?.label : 'all statuses'}
                </span>
              )}
            </span>
            {mfTotals && (
              <div className="card-head-actions">
                <select
                  className="ov-ward-status-select"
                  value={wardStatusFilter}
                  onChange={(e) => setWardStatusFilter(e.target.value as MfStatusKey | '')}
                  title="Filter ward downloads by status"
                >
                  <option value="">Status: All</option>
                  {MF_TILES.map(({ key, label }) => (
                    <option key={key} value={key}>
                      {label}
                    </option>
                  ))}
                </select>
              </div>
            )}
          </div>
          {!mfTotals || wardTypes.length === 0 ? (
            <div className="empty">
              <p>Import the office's Monitoring Format workbook (Works List page) to see ward/circle-wise totals here.</p>
            </div>
          ) : (
            <div className="ov-mf-body">
              <div className="ov-mf-donut-col">
                <div className="ov-mf-donut-wrap">
                  <svg viewBox="0 0 120 120" className="ov-donut ov-mf-donut">
                    <circle cx="60" cy="60" r={DONUT_R} className="ov-donut-track" />
                    {(() => {
                      const wardTotal = wardTypes.reduce((s, w) => s + w.row.totalWorks.no, 0)
                      let cursor = 0
                      return wardTypes.map(({ row, color }) => {
                        const share = wardTotal > 0 ? row.totalWorks.no / wardTotal : 0
                        if (share <= 0) return null
                        const segLen = share * DONUT_C
                        const gap = 2
                        const dash = `${Math.max(segLen - gap, 0)} ${DONUT_C - segLen + gap}`
                        const offset = -cursor
                        cursor += segLen
                        return (
                          <circle
                            key={row.itemType}
                            cx="60"
                            cy="60"
                            r={DONUT_R}
                            className="ov-mf-donut-seg"
                            style={{ stroke: color, strokeDasharray: dash, strokeDashoffset: offset }}
                          />
                        )
                      })
                    })()}
                  </svg>
                  <div className="ov-donut-text">
                    <b>{wardTypes.reduce((s, w) => s + w.row.totalWorks.no, 0)}</b>
                    <span>Total Works</span>
                  </div>
                </div>

                <ul className="ov-donut-legend-list">
                  {wardTypes.map(({ row, color }) => (
                    <li key={row.itemType}>
                      <i className="ov-donut-legend-dot" style={{ background: color }} />
                      <span>{row.itemType}</span>
                      <b>{row.totalWorks.no}</b>
                    </li>
                  ))}
                </ul>
              </div>

              <div className="ov-mf-tiles">
                {wardTypes.map(({ row, color }) => {
                  const pct =
                    mfTotals.totalWorks.no > 0 ? Math.round((row.totalWorks.no / mfTotals.totalWorks.no) * 100) : 0
                  const isExporting = exportingWard === row.itemType
                  const isExported = exportedWard === row.itemType
                  return (
                    <button
                      key={row.itemType}
                      type="button"
                      className={`ov-mf-tile${isExporting ? ' ov-mf-tile-busy' : ''}${isExported ? ' ov-mf-tile-done' : ''}`}
                      style={{ ['--tile-color' as string]: color }}
                      disabled={row.totalWorks.no === 0 || isExporting}
                      title={
                        row.totalWorks.no > 0
                          ? `Download Excel of ${row.itemType}${
                              wardStatusFilter ? ` — ${MF_TILES.find((t) => t.key === wardStatusFilter)?.label}` : ''
                            } works`
                          : undefined
                      }
                      onClick={() => handleExportWardTile(row.itemType)}
                    >
                      <span className="ov-mf-tile-ic">{isExporting ? <IconRefresh className="ov-mf-tile-spin" /> : isExported ? <IconCheck /> : null}</span>
                      <span className="ov-mf-tile-main">
                        <span className="ov-mf-tile-label">
                          {row.itemType}
                          {isExporting && <span className="ov-mf-tile-status"> · Downloading…</span>}
                          {isExported && <span className="ov-mf-tile-status ov-mf-tile-status-done"> · Downloaded</span>}
                        </span>
                        <span className="ov-mf-tile-nums">
                          <b className="ov-mf-tile-no">{row.totalWorks.no}</b>
                          <span className="ov-mf-tile-pct">{pct}%</span>
                        </span>
                        <span className="ov-mf-tile-amt">Rs. {indianDigitGroups(row.totalWorks.amt)} L</span>
                      </span>
                    </button>
                  )
                })}
              </div>
            </div>
          )}
        </div>
      </div>

      <div className="ov-grid">
        <div className="card ov-card ov-mb">
          <div className="card-head">
            <h3>MB Scrutiny</h3>
            <button className="ghost ov-new-btn" onClick={() => onNavigate('mbScrutiny')}>
              <IconPlus /> Add MB
            </button>
          </div>
          {recentMb.length === 0 ? (
            <div className="empty">
              <p>No MBs received yet.</p>
            </div>
          ) : (
            <ul className="ov-mb-list">
              {recentMb.map((m) => (
                <li key={m.id}>
                  <span className="ov-mb-name">
                    {m.agencyName}
                    <span className="ov-mb-sub">MB No. {m.mbNo}</span>
                  </span>
                  <span className={`ov-status-chip ${m.done ? 'ov-status-done' : 'ov-status-progress'}`}>
                    {m.done ? 'Completed' : 'In Progress'}
                  </span>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="card ov-card ov-reminder">
          <div className="card-head">
            <h3>Reminders</h3>
          </div>
          {nextDeadline ? (
            <>
              <p className="ov-reminder-title">{nextDeadline.text}</p>
              <p className="ov-reminder-date">
                {nextDeadline.kind === 'tender' ? nextDeadline.dueLabel : `Due ${nextDeadline.dueLabel}`}
              </p>
              <button className="primary ov-reminder-btn" onClick={() => onNavigate(deadlineTab(nextDeadline.kind))}>
                Open
              </button>
            </>
          ) : (
            <div className="empty">
              <p>No upcoming deadlines.</p>
            </div>
          )}
        </div>

        <div className="card ov-card ov-tasks">
          <div className="card-head">
            <h3>Upcoming</h3>
            <button className="ghost ov-new-btn" onClick={() => onNavigate('todo')}>
              <IconPlus /> New
            </button>
          </div>
          {upcoming.length === 0 ? (
            <div className="empty">
              <p>Nothing due yet.</p>
            </div>
          ) : (
            <ul className="ov-task-list">
              {upcoming.map((d) => (
                <li key={`${d.kind}-${d.id}`}>
                  <span className={`ov-task-ic ov-task-ic-${d.kind}`}>
                    {d.kind === 'mb' ? <IconEye /> : d.kind === 'tender' ? <IconSearch /> : <IconChecklist />}
                  </span>
                  <span className="ov-task-body">
                    <span className="ov-task-text">{d.text}</span>
                    <span className="ov-task-date">{d.kind === 'tender' ? d.dueLabel : `Due date: ${d.dueLabel}`}</span>
                  </span>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="card ov-card ov-progress">
          <div className="card-head">
            <h3>Task Progress</h3>
          </div>
          <div className="ov-donut-wrap">
            <svg viewBox="0 0 120 120" className="ov-donut">
              <circle cx="60" cy="60" r="50" className="ov-donut-track" />
              <circle
                cx="60"
                cy="60"
                r="50"
                className="ov-donut-fill"
                strokeDasharray={`${(progressPct / 100) * 314.16} 314.16`}
              />
            </svg>
            <div className="ov-donut-text">
              <b>{progressPct}%</b>
              <span>Completed</span>
            </div>
          </div>
          <div className="ov-progress-bar" style={{ ['--done-pct' as string]: `${progressPct}%` }} />
          <div className="ov-donut-legend">
            <span>
              <i className="ov-dot ov-dot-done" /> Completed ({completedCount})
            </span>
            <span>
              <i className="ov-dot ov-dot-pending" /> Pending ({pendingCount})
            </span>
          </div>
        </div>

        <button className="ov-card ov-promo" onClick={() => onNavigate('cementSteel')}>
          <div className="ov-promo-head">
            <IconDoc />
            {cementSteelHasNew && <span className="ov-promo-badge">New rates</span>}
          </div>
          <p className="ov-promo-title">Cement &amp; Steel Rates</p>
          <p className="ov-promo-sub">Check the latest circular rates.</p>
          <span className="ov-promo-btn">Open</span>
        </button>

        <WeatherCard />
      </div>
    </div>
  )
}

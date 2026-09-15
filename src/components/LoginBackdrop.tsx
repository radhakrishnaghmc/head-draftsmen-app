import { Fragment } from 'react'
import appLogo from '../assets/app-logo.png'
import {
  IconGrid,
  IconTable,
  IconCalendar,
  IconBell,
  IconClipboard,
  IconWarn,
  IconPrint,
  IconDownload,
  IconChecklist,
  IconEye,
  IconSearch,
  IconCheck,
  IconTools,
  IconDoc,
  IconSettings
} from './Icons'

// Static stand-in for Sidebar's real nav list — labels only, no badges/counts
// (those are per-office live data this component never sees).
const NAV_ITEMS: { label: string; icon: JSX.Element; on?: boolean; section?: string }[] = [
  { label: 'Dashboard', icon: <IconGrid />, on: true, section: 'Overview' },
  { label: 'Works List', icon: <IconTable /> },
  { label: 'Calendar', icon: <IconCalendar /> },
  { label: 'Intimation', icon: <IconBell /> },
  { label: 'Agreement and Work order', icon: <IconClipboard /> },
  { label: 'Issue Notices', icon: <IconWarn /> },
  { label: 'Issue Documents', icon: <IconPrint /> },
  { label: 'BOQ, Schedule A, Deviation & Material Quantity', icon: <IconDownload /> },
  { label: 'To Do List', icon: <IconChecklist /> },
  { label: 'MB Scrutiny list', icon: <IconEye /> },
  { label: 'Search Tender', icon: <IconSearch /> },
  { label: 'Give Technical Sanction (Experimental)', icon: <IconCheck /> },
  { label: 'Tools', icon: <IconTools />, section: 'Utilities' },
  { label: 'Cement & Steel Rates', icon: <IconDoc /> },
  { label: 'Settings', icon: <IconSettings />, section: 'System' }
]

/**
 * Purely decorative recreation of the real app shell (Sidebar +
 * Dashboard), shown blurred behind the login card (see AuthGate) so the
 * login screen reads as "this app" instead of a blank void. Static markup
 * only — same CSS classes as Sidebar/OverviewDashboard so theming (Default/
 * Aurora/Paper) matches exactly, but no live data, no IPC calls, and nothing
 * interactive (the wrapping .login-backdrop-app sets pointer-events: none).
 */
export default function LoginBackdrop() {
  return (
    <div className="shell" aria-hidden="true">
      <aside className="sidebar">
        <div className="side-brand">
          <img src={appLogo} alt="" className="logo" />
          <div className="side-brand-text">
            <h1>Head Draughtsman</h1>
            <p>Corporation · Zone · Circle</p>
          </div>
        </div>
        <nav className="side-nav">
          {NAV_ITEMS.map((it, i) => (
            <Fragment key={i}>
              {it.section ? <div className="nav-section-label">{it.section}</div> : null}
              <div className={`nav-item ${it.on ? 'on' : ''}`}>
                <span className="nav-ic">{it.icon}</span>
                <span className="nav-text">
                  <span className="nav-label">{it.label}</span>
                </span>
              </div>
            </Fragment>
          ))}
        </nav>
      </aside>

      <main className="workspace">
        <section className="page wide">
          <div className="page-head">
            <div className="page-ic green">
              <IconGrid />
            </div>
            <div className="page-head-text">
              <h1>Dashboard</h1>
            </div>
          </div>
          <div className="overview-dashboard">
            <div className="ov-stat-row">
              <div className="ov-stat ov-stat-hero">
                <span className="ov-stat-top">Total works</span>
                <b>—</b>
                <span className="ov-stat-sub">Sign in to load your office</span>
              </div>
              <div className="ov-stat">
                <span className="ov-stat-top">Pending items</span>
                <b>—</b>
                <span className="ov-stat-sub">To Do &amp; MB Scrutiny open</span>
              </div>
              <div className="ov-stat">
                <span className="ov-stat-top">Completed</span>
                <b>—</b>
                <span className="ov-stat-sub">— of — done</span>
              </div>
              <div className="ov-stat">
                <span className="ov-stat-top">Next deadline</span>
                <b>—</b>
                <span className="ov-stat-sub">Nothing scheduled</span>
              </div>
            </div>
          </div>
        </section>
      </main>
    </div>
  )
}

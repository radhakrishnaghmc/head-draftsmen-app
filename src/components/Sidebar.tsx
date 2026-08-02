import { useEffect, useRef, useState } from 'react'
import { api } from '../ipc'
import appLogo from '../assets/app-logo.png'
import { type Office, isOfficeReady, normalizeOffice } from '../office'
import { CORPORATIONS, zonesOf, circlesOf } from '../zoneCircleDirectory'
import {
  IconTable,
  IconPrint,
  IconWarn,
  IconCheck,
  IconSearch,
  IconCalendar,
  IconDownload,
  IconChecklist,
  IconClipboard,
  IconEye,
  IconRefresh,
  IconBell,
  IconTools,
  IconWhatsApp
} from './Icons'

export type TabKey =
  | 'dashboard'
  | 'data'
  | 'printDoc'
  | 'estimateWorkspace'
  | 'techSanction'
  | 'intimation'
  | 'workOrder'
  | 'mbScrutiny'
  | 'search'
  | 'todo'
  | 'tools'

interface Props {
  active: TabKey
  onSelect: (t: TabKey) => void
  tableCount: number
  unresolved: number
  createdDocCount: number
  /** The chosen office (Corporation/Zone/Circle) — shown under the app name; also picked here and in the Works List. */
  office: Office
  /** Change the office — lets the sidebar's own picker set it, same as the Works List one. */
  onOfficeChange: (office: Office) => void
}

interface Item {
  key: TabKey
  label: string
  icon: JSX.Element
  badge?: string | number
  warn?: boolean
  ok?: boolean
  tone?: 'green' | 'sky' | 'rose' | 'amber' | 'teal'
}

const UPDATE_STATUS_TEXT: Record<string, string> = {
  'update-available': 'Update found — downloading…',
  'up-to-date': "You're up to date",
  error: "Couldn't check for updates",
  'dev-mode': 'Not available in development'
}

export default function Sidebar(props: Props) {
  // The office picker opens as a popup (kept out of the way of the clean brand
  // header); it opens itself when no office is chosen yet.
  const [officeOpen, setOfficeOpen] = useState(!isOfficeReady(props.office))
  const [version, setVersion] = useState<string | null>(null)
  useEffect(() => {
    api.getAppVersion().then(setVersion)
  }, [])

  const [checkingUpdate, setCheckingUpdate] = useState(false)
  const [updateStatus, setUpdateStatus] = useState<string | null>(null)
  const clearStatusTimer = useRef<ReturnType<typeof setTimeout> | null>(null)

  async function checkForUpdates() {
    if (checkingUpdate) return
    setCheckingUpdate(true)
    setUpdateStatus(null)
    if (clearStatusTimer.current) clearTimeout(clearStatusTimer.current)
    try {
      const result = await api.checkForUpdates()
      setUpdateStatus(UPDATE_STATUS_TEXT[result] ?? null)
    } catch {
      setUpdateStatus(UPDATE_STATUS_TEXT.error)
    } finally {
      setCheckingUpdate(false)
      clearStatusTimer.current = setTimeout(() => setUpdateStatus(null), 4000)
    }
  }

  const items: Item[] = [
    {
      key: 'data',
      label: 'Works List',
      icon: <IconTable />,
      badge: props.tableCount || undefined,
      warn: props.unresolved > 0,
      tone: 'green'
    },
    {
      key: 'dashboard',
      label: 'Calendar',
      icon: <IconCalendar />
    },
    {
      key: 'intimation',
      label: 'Intimation',
      icon: <IconBell />,
      tone: 'rose'
    },
    {
      key: 'workOrder',
      label: 'Agreement and Work order',
      icon: <IconClipboard />
    },
    {
      key: 'printDoc',
      label: 'Issue Documents',
      icon: <IconPrint />,
      badge: props.createdDocCount || undefined,
      tone: 'sky'
    },
    {
      key: 'estimateWorkspace',
      label: 'BOQ, Schedule A, Deviation & Material Quantity',
      icon: <IconDownload />,
      tone: 'amber'
    },
    {
      key: 'todo',
      label: 'To Do List',
      icon: <IconChecklist />
    },
    {
      key: 'mbScrutiny',
      label: 'MB Scrutiny list',
      icon: <IconEye />
    },
    {
      key: 'search',
      label: 'Search Tender',
      icon: <IconSearch />,
      tone: 'green'
    },
    {
      key: 'techSanction',
      label: 'Give Technical Sanction (Experimental)',
      icon: <IconCheck />
    },
    {
      key: 'tools',
      label: 'Tools',
      icon: <IconTools />,
      tone: 'teal'
    }
  ]

  return (
    <aside className="sidebar">
      <div className="side-brand">
        <img src={appLogo} alt="" className="logo" />
        <div className="side-brand-text">
          <h1>Head Draughtsman</h1>
          {isOfficeReady(props.office) ? (
            <p>
              {[props.office.corporation, props.office.zone, props.office.circle, props.office.circleNumber]
                .filter(Boolean)
                .join(' · ')}
            </p>
          ) : (
            <p className="warn">Select your office</p>
          )}
          <button className="side-office-change" onClick={() => setOfficeOpen(true)}>
            Change office
          </button>
        </div>
      </div>

      <nav className="side-nav">
        {items.map((it, i) => (
          <button
            key={it.key}
            className={`nav-item ${props.active === it.key ? 'on' : ''} ${it.tone ? `tone-${it.tone}` : ''}`}
            onClick={() => props.onSelect(it.key)}
          >
            <span className="nav-step">{i + 1}</span>
            <span className="nav-ic">{it.icon}</span>
            <span className="nav-text">
              <span className="nav-label">{it.label}</span>
            </span>
            {it.warn ? (
              <span className="nav-badge warn">
                <IconWarn />
              </span>
            ) : it.ok ? (
              <span className="nav-badge ok">
                <IconCheck />
              </span>
            ) : it.badge != null ? (
              <span className="nav-badge">{it.badge}</span>
            ) : null}
          </button>
        ))}
      </nav>

      <div className="side-foot">
        <div className="side-foot-row">
          <span className="side-credits">
            App developed by Radhakrishna, HD{version ? ` · v${version}` : ''}
            <button
              type="button"
              className="wa-contact"
              title="Contact Admin on WhatsApp"
              onClick={() => api.openPath('https://wa.me/919063836115?text=Hello')}
            >
              <IconWhatsApp /> Contact Admin on WhatsApp
            </button>
          </span>
          <button
            className={`side-update-btn ${checkingUpdate ? 'spinning' : ''}`}
            title="Check for updates"
            onClick={checkForUpdates}
            disabled={checkingUpdate}
          >
            <IconRefresh />
          </button>
        </div>
        {updateStatus && <span className="side-update-status">{updateStatus}</span>}
      </div>

      {officeOpen && (
        <div className="editor-overlay" onClick={() => setOfficeOpen(false)}>
          <div className="confirm-modal office-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
            <h3>Your office</h3>
            <div className="office-fields">
              <div className="office-field">
                <span className="office-field-label">Corporation</span>
                <div className="office-chips">
                  {CORPORATIONS.map((c) => (
                    <button
                      key={c.name}
                      className={`office-chip ${props.office.corporation === c.name ? 'on' : ''}`}
                      title={c.fullName}
                      onClick={() =>
                        props.onOfficeChange(
                          normalizeOffice({ corporation: c.name === props.office.corporation ? undefined : c.name })
                        )
                      }
                    >
                      {c.name}
                    </button>
                  ))}
                </div>
              </div>

              {props.office.corporation && (
                <div className="office-field">
                  <span className="office-field-label">Zone</span>
                  <div className="office-chips">
                    {zonesOf(props.office.corporation).length === 0 ? (
                      <span className="office-empty">No zones listed for {props.office.corporation} yet.</span>
                    ) : (
                      zonesOf(props.office.corporation).map((z) => (
                        <button
                          key={z}
                          className={`office-chip ${props.office.zone === z ? 'on' : ''}`}
                          onClick={() =>
                            props.onOfficeChange(
                              normalizeOffice({
                                corporation: props.office.corporation,
                                zone: z === props.office.zone ? undefined : z
                              })
                            )
                          }
                        >
                          {z}
                        </button>
                      ))
                    )}
                  </div>
                </div>
              )}

              {props.office.zone && circlesOf(props.office.corporation, props.office.zone).length > 0 && (
                <div className="office-field">
                  <span className="office-field-label">Circle · leave blank for a zonal office</span>
                  <div className="office-chips">
                    {circlesOf(props.office.corporation, props.office.zone).map((c) => (
                      <button
                        key={c.circle}
                        className={`office-chip ${props.office.circle === c.circle ? 'on' : ''}`}
                        onClick={() =>
                          props.onOfficeChange(
                            normalizeOffice({
                              corporation: props.office.corporation,
                              zone: props.office.zone,
                              circle: c.circle === props.office.circle ? undefined : c.circle
                            })
                          )
                        }
                      >
                        {c.circle}
                        <span className="office-chip-cno">{c.cno}</span>
                      </button>
                    ))}
                  </div>
                </div>
              )}
            </div>
            <div className="confirm-actions">
              <button className="primary" onClick={() => setOfficeOpen(false)}>
                Done
              </button>
            </div>
          </div>
        </div>
      )}
    </aside>
  )
}

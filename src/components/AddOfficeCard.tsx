import { useState } from 'react'
import { createPortal } from 'react-dom'
import { CORPORATIONS, zonesOf, circlesOf } from '../zoneCircleDirectory'
import { normalizeOffice, type Office } from '../office'
import { MAX_ACCOUNT_OFFICES, sameAccountOffice, type AccountOffice } from '@core/signupRequest'
import { IconWarn, IconUser } from './Icons'

interface Props {
  /** The offices on this login (the sign-up office first). */
  offices: AccountOffice[]
  /** The office currently open in the app. */
  active: Office
  onAdd: (office: AccountOffice) => Promise<string>
  onRemove: (office: AccountOffice) => Promise<string>
}

const label = (o: AccountOffice) => [o.corporation, o.zone, o.circle || 'Zonal office'].filter(Boolean).join(' · ')

/**
 * Settings → Add office. A login holds at most MAX_ACCOUNT_OFFICES offices; the office details are picked the same way
 * the sidebar's office switch used to (Corporation → Zone → Circle chips, circle blank = zonal office).
 */
export default function AddOfficeCard({ offices, active, onAdd, onRemove }: Props) {
  const [adding, setAdding] = useState(false)
  const [pick, setPick] = useState<Office>({})
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const full = offices.length >= MAX_ACCOUNT_OFFICES

  async function save() {
    setBusy(true)
    const problem = await onAdd({
      corporation: pick.corporation ?? '',
      zone: pick.zone ?? '',
      circle: pick.circle ?? ''
    })
    setBusy(false)
    if (problem) setError(problem)
    else {
      setAdding(false)
      setPick({})
      setError(null)
    }
  }

  async function remove(o: AccountOffice) {
    const first = offices[0] && sameAccountOffice(offices[0], o)
    const ok = window.confirm(
      `Delete the office "${label(o)}" from this login?` +
        (first ? "\n\nThis is the office you signed up with — your other office becomes the login's main office." : '')
    )
    if (!ok) return
    setBusy(true)
    const problem = await onRemove(o)
    setBusy(false)
    setError(problem || null)
  }

  const zones = zonesOf(pick.corporation)
  const circles = circlesOf(pick.corporation, pick.zone)

  return (
    <div className="card">
      <div className="card-head">
        <div className="head-ic">
          <IconUser />
        </div>
        <div className="titles">
          <h2>Offices on this login</h2>
          <p className="sub">
            {offices.length} of {MAX_ACCOUNT_OFFICES} offices. Each office keeps its own Works List and documents.
          </p>
        </div>
      </div>

      {full && (
        <div className="notice error">
          <IconWarn />
          Only {MAX_ACCOUNT_OFFICES} offices per login are allowed. To add another office, delete one of the two.
        </div>
      )}
      {error && !adding && (
        <div className="notice error">
          <IconWarn />
          {error}
        </div>
      )}

      <ul className="active-devices-list">
        {offices.map((o) => {
          const open = sameAccountOffice(o, { corporation: active.corporation ?? '', zone: active.zone ?? '', circle: active.circle ?? '' })
          return (
            <li key={label(o)} className="active-device-row">
              <div className="active-device-info">
                <span className="active-device-label">
                  {label(o)}
                  {open && <span className="active-device-tag">Open now</span>}
                </span>
                <span className="active-device-meta">
                  {o === offices[0] ? 'Main office (signed up with)' : 'Added office'}
                </span>
              </div>
              <button type="button" className="active-device-logout" disabled={busy} onClick={() => void remove(o)}>
                Delete
              </button>
            </li>
          )
        })}
      </ul>

      <div className="confirm-actions">
        <button
          className="primary"
          disabled={busy}
          onClick={() => {
            setError(null)
            if (full)
              setError(
                `Only ${MAX_ACCOUNT_OFFICES} offices per login are allowed. To add another, delete one of the two.`
              )
            else setAdding(true)
          }}
        >
          Add office
        </button>
      </div>

      {adding &&
        createPortal(
          <div className="editor-overlay">
            <div className="confirm-modal office-modal" role="dialog" aria-modal="true">
              <h3>Add office</h3>
              <div className="office-fields">
                <div className="office-field">
                  <span className="office-field-label">Corporation</span>
                  <div className="office-chips">
                    {CORPORATIONS.map((c) => (
                      <button
                        key={c.name}
                        className={`office-chip ${pick.corporation === c.name ? 'on' : ''}`}
                        title={c.fullName}
                        onClick={() =>
                          setPick(normalizeOffice({ corporation: c.name === pick.corporation ? undefined : c.name }))
                        }
                      >
                        {c.name}
                      </button>
                    ))}
                  </div>
                </div>

                {pick.corporation && (
                  <div className="office-field">
                    <span className="office-field-label">Zone</span>
                    <div className="office-chips">
                      {zones.length === 0 ? (
                        <span className="office-empty">No zones listed for {pick.corporation} yet.</span>
                      ) : (
                        zones.map((z) => (
                          <button
                            key={z}
                            className={`office-chip ${pick.zone === z ? 'on' : ''}`}
                            onClick={() =>
                              setPick(
                                normalizeOffice({
                                  corporation: pick.corporation,
                                  zone: z === pick.zone ? undefined : z
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

                {pick.zone && circles.length > 0 && (
                  <div className="office-field">
                    <span className="office-field-label">Circle · leave blank for a zonal office</span>
                    <div className="office-chips">
                      {circles.map((c) => (
                        <button
                          key={c.circle}
                          className={`office-chip ${pick.circle === c.circle ? 'on' : ''}`}
                          onClick={() =>
                            setPick(
                              normalizeOffice({
                                corporation: pick.corporation,
                                zone: pick.zone,
                                circle: c.circle === pick.circle ? undefined : c.circle
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
              {error && (
                <div className="notice error">
                  <IconWarn />
                  {error}
                </div>
              )}
              <div className="confirm-actions">
                <button
                  className="ghost"
                  disabled={busy}
                  onClick={() => {
                    setAdding(false)
                    setPick({})
                    setError(null)
                  }}
                >
                  Cancel
                </button>
                <button
                  className="primary"
                  disabled={busy || !pick.corporation || !pick.zone}
                  onClick={() => void save()}
                >
                  {busy ? 'Adding…' : 'Add office'}
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
    </div>
  )
}

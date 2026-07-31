import { useState } from 'react'
import { CORPORATIONS, zonesOf, circlesOf } from '../zoneCircleDirectory'
import { type Office, normalizeOffice, isOfficeReady } from '../office'
import { IconCheck, IconWarn } from './Icons'

interface Props {
  office: Office
  onChange: (office: Office) => void
}

/**
 * Cascading Corporation → Zone → Circle picker shown under the app name in the
 * sidebar. The choice here (not the login) is what documents and Works List
 * validation are prepared for. Collapsed to a one-line summary once set; opens
 * to the button rows to change it (and opens itself when nothing's chosen yet).
 */
export default function OfficeSelector({ office, onChange }: Props) {
  const ready = isOfficeReady(office)
  // Open by default until an office is chosen — it's required before documents
  // can be prepared for the right circle.
  const [open, setOpen] = useState(!ready)

  const zones = zonesOf(office.corporation)
  const circles = circlesOf(office.corporation, office.zone)

  // Selecting higher in the chain clears everything below it.
  const pickCorporation = (corporation: string) =>
    onChange(normalizeOffice({ corporation: corporation === office.corporation ? undefined : corporation }))
  const pickZone = (zone: string) =>
    onChange(normalizeOffice({ corporation: office.corporation, zone: zone === office.zone ? undefined : zone }))
  const pickCircle = (circle: string) =>
    onChange(
      normalizeOffice({
        corporation: office.corporation,
        zone: office.zone,
        circle: circle === office.circle ? undefined : circle
      })
    )

  // Corporation · Zone · Circle · Circle number — all four, each its own segment.
  const summary = [office.corporation, office.zone, office.circle, office.circleNumber]
    .filter(Boolean)
    .join(' · ')

  return (
    <div className={`office-picker ${ready ? '' : 'unset'}`}>
      <button className="office-summary" onClick={() => setOpen((v) => !v)} aria-expanded={open}>
        {ready ? (
          <span className="office-summary-text">
            <IconCheck /> {summary}
          </span>
        ) : (
          <span className="office-summary-text warn">
            <IconWarn /> {summary || 'Select your office'}
          </span>
        )}
        <span className="office-summary-toggle">{open ? 'Done' : 'Change your Office'}</span>
      </button>

      {open && (
        <div className="office-fields">
          <div className="office-field">
            <span className="office-field-label">Corporation</span>
            <div className="office-chips">
              {CORPORATIONS.map((c) => (
                <button
                  key={c.name}
                  className={`office-chip ${office.corporation === c.name ? 'on' : ''}`}
                  onClick={() => pickCorporation(c.name)}
                  title={c.fullName}
                >
                  {c.name}
                </button>
              ))}
            </div>
          </div>

          {office.corporation && (
            <div className="office-field">
              <span className="office-field-label">Zone</span>
              <div className="office-chips">
                {zones.length === 0 ? (
                  <span className="office-empty">No zones listed for {office.corporation} yet.</span>
                ) : (
                  zones.map((z) => (
                    <button
                      key={z}
                      className={`office-chip ${office.zone === z ? 'on' : ''}`}
                      onClick={() => pickZone(z)}
                    >
                      {z}
                    </button>
                  ))
                )}
              </div>
            </div>
          )}

          {office.zone && circles.length > 0 && (
            <div className="office-field">
              <span className="office-field-label">Circle · leave blank for a zonal office</span>
              <div className="office-chips">
                {circles.map((c) => (
                  <button
                    key={c.circle}
                    className={`office-chip ${office.circle === c.circle ? 'on' : ''}`}
                    onClick={() => pickCircle(c.circle)}
                  >
                    {c.circle}
                    <span className="office-chip-cno">{c.cno}</span>
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  )
}

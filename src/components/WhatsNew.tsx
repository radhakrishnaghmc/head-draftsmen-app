import type { ChangelogEntry } from '@core/changelog'
import { IconBolt, IconCheck } from './Icons'
import { closeOnBackdropMouseDown } from '../overlayClose'

interface Props {
  /** Versions to announce, newest first (already filtered to what's new since last seen). */
  entries: ChangelogEntry[]
  onClose: () => void
}

/**
 * "What's New" modal, shown once on the first launch after an update. Lists the
 * changes for every version newer than the one last seen on this machine.
 */
export default function WhatsNew({ entries, onClose }: Props) {
  if (entries.length === 0) return null
  const heading = entries.length === 1 ? `What's new in v${entries[0].version}` : "What's new"

  return (
    <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(onClose)}>
      <div
        className="confirm-modal whatsnew-modal"
        role="dialog"
        aria-modal="true"
        aria-label={heading}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="confirm-ic whatsnew-ic">
          <IconBolt />
        </div>
        <h3>{heading}</h3>
        <p className="whatsnew-sub">Here's what changed since you last used the app.</p>

        <div className="whatsnew-body">
          {entries.map((entry) => (
            <div key={entry.version} className="whatsnew-version">
              {entries.length > 1 && <div className="whatsnew-version-tag">v{entry.version}</div>}
              <ul className="whatsnew-list">
                {entry.changes.map((change, i) => (
                  <li key={i}>
                    <IconCheck />
                    <span>{change}</span>
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>

        <div className="confirm-actions whatsnew-actions">
          <button className="primary" onClick={onClose} autoFocus>
            Got it
          </button>
        </div>
      </div>
    </div>
  )
}

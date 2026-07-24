import { IconWarn } from './Icons'
import type { Collision, CollisionResolution } from '@core/types'

interface Props {
  collisions: Collision[]
  resolution: CollisionResolution
  onResolve: (column: string, source: string) => void
}

export default function CollisionPanel({ collisions, resolution, onResolve }: Props) {
  if (collisions.length === 0) return null

  return (
    <section className="card full-width">
      <div className="card-head">
        <div className="step-num" style={{ color: 'var(--warn)', background: 'var(--warn-bg)', borderColor: '#fde68a' }}>
          <IconWarn />
        </div>
        <div className="titles">
          <h2>Resolve column collisions</h2>
          <p className="sub">The same column exists in multiple files — choose the source of truth</p>
        </div>
      </div>

      <div className="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>Column</th>
              <th>Appears in</th>
              <th>Use source</th>
            </tr>
          </thead>
          <tbody>
            {collisions.map((c) => (
              <tr key={c.column}>
                <td>
                  <code>{c.column}</code>
                </td>
                <td className="mono">{c.sources.join('  ·  ')}</td>
                <td>
                  <select
                    value={resolution[c.column] ?? ''}
                    onChange={(e) => onResolve(c.column, e.target.value)}
                  >
                    <option value="" disabled>
                      Select…
                    </option>
                    {c.sources.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

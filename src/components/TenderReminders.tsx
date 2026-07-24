import { IconRefresh, IconTrash, IconChecklist } from './Icons'
import type { TenderReminder } from '@core/types'

interface Props {
  reminders: TenderReminder[]
  onDelete: (id: string) => void
  onRefresh: (id: string) => void
  refreshingId: string | null
}

/**
 * Reminders for tenders just issued a notice — created automatically from the
 * Calendar section's "Issue tender notice" flow, matched against the live
 * e-procurement portal for their bid-closing time.
 */
export default function TenderReminders({ reminders, onDelete, onRefresh, refreshingId }: Props) {
  if (reminders.length === 0) return null

  return (
    <section className="card">
      <div className="card-head">
        <div className="head-ic">
          <IconChecklist />
        </div>
        <div className="titles">
          <h2>Tender Reminders</h2>
          <p className="sub">From tender notices you've issued — bid-closing times pulled live.</p>
        </div>
      </div>
      <ul className="todo-list boq-entries">
        {reminders.map((r) => (
          <li key={r.id} className="todo-item">
            <div className="todo-body">
              {r.items.length > 1 ? (
                <>
                  <span className="todo-text">{r.nitNo}</span>
                  <ul className="reminder-works">
                    {r.items.map((it, i) => (
                      <li key={i}>
                        <span className="reminder-work-name">{it.workName || 'Work'}</span>
                        {it.tenderId && <span> · Tender ID: {it.tenderId}</span>}
                        {it.bidClosing && <span> · Bid closing: {it.bidClosing}</span>}
                      </li>
                    ))}
                  </ul>
                </>
              ) : (
                <>
                  <span className="todo-text">{r.items[0]?.workName || r.nitNo}</span>
                  <span className="todo-dates">
                    {r.status === 'found' &&
                      r.items[0]?.tenderId &&
                      `Tender ID: ${r.items[0].tenderId}${r.items[0].bidClosing ? ` · Bid closing: ${r.items[0].bidClosing}` : ''}`}
                    {r.status === 'found' && !r.items[0]?.tenderId && r.items[0]?.bidClosing && `Bid closing: ${r.items[0].bidClosing}`}
                    {r.status === 'pending' && 'Looking up bid closing time…'}
                    {r.status === 'not-found' &&
                      'Not found on the portal yet — it may take a while to be listed.'}
                  </span>
                </>
              )}
            </div>
            <button
              className="ghost"
              title="Re-check the portal"
              onClick={() => onRefresh(r.id)}
              disabled={refreshingId === r.id}
            >
              <IconRefresh />
            </button>
            <button className="danger-ghost" title="Delete reminder" onClick={() => onDelete(r.id)}>
              <IconTrash />
            </button>
          </li>
        ))}
      </ul>
    </section>
  )
}

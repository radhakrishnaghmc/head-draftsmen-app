import { useState } from 'react'
import { IconRefresh } from './Icons'

/** "Refresh" for an issued-documents board: re-reads the matching Works List rows and reports how many were found. */
export default function RefreshFromWorksButton({ hasWorks, onRefresh }: { hasWorks: boolean; onRefresh: () => { found: number; total: number } }) {
  const [note, setNote] = useState('')
  return (
    <>
      <button
        className="ghost"
        title="Refresh Tender ID, NIT, agency and estimate amount from the Works List, matched by name of the work"
        onClick={() => {
          if (!hasWorks) return setNote('No Works List loaded.')
          const { found, total } = onRefresh()
          setNote(`${found} of ${total} refreshed from the Works List${found < total ? ` (${total - found} not found by name)` : ''}.`)
        }}
      >
        <IconRefresh /> Refresh
      </button>
      {note && <span className="muted">{note}</span>}
    </>
  )
}

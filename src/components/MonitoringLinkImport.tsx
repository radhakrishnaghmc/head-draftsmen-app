import { useState } from 'react'
import { IconLink, IconWarn, IconCheck } from './Icons'

export interface MonitoringMergeSummary {
  matched: number
  added: number
  filled: number
  sheetName: string
}

interface Props {
  onImport: (url: string) => Promise<MonitoringMergeSummary>
}

/**
 * A monitoring format is one workbook with a sheet per circle (grouped under
 * zone tabs). Pasting its link picks out the one sheet matching the logged-in
 * Circle and merges it into the existing Works List — filling in blank
 * columns for works already there (matched by Wincode or Name of the work)
 * and adding any work that isn't in the database yet. Unlike the plain
 * Works List import above, this never replaces existing data.
 */
export default function MonitoringLinkImport({ onImport }: Props) {
  const [url, setUrl] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [done, setDone] = useState<string | null>(null)

  async function run() {
    const trimmed = url.trim()
    if (!trimmed || loading) return
    setLoading(true)
    setError(null)
    setDone(null)
    try {
      const { matched, added, filled, sheetName } = await onImport(trimmed)
      setDone(
        `Merged the "${sheetName}" sheet: ${matched} work${matched === 1 ? '' : 's'} matched ` +
          `(${filled} cell${filled === 1 ? '' : 's'} filled in), ${added} new work${added === 1 ? '' : 's'} added.`
      )
      setUrl('')
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="card link-import">
      <div className="link-import-row">
        <IconLink />
        <input
          value={url}
          placeholder="Paste a monitoring format link…"
          onChange={(e) => {
            setUrl(e.target.value)
            setError(null)
            setDone(null)
          }}
          onKeyDown={(e) => e.key === 'Enter' && run()}
          disabled={loading}
        />
        <button className="primary" onClick={run} disabled={loading || !url.trim()}>
          {loading ? 'Merging…' : 'Merge from Monitoring Format'}
        </button>
      </div>
      <p className="hint">
        Picks the sheet matching your Circle and fills in blank columns for matching works (by
        Wincode or Name of the work), adding any work not already in the database. Existing data is
        never overwritten.
      </p>
      {error && (
        <div className="notice error">
          <IconWarn />
          {error}
        </div>
      )}
      {done && (
        <div className="notice ok">
          <IconCheck />
          {done}
        </div>
      )}
    </div>
  )
}

import { useState } from 'react'
import { IconLink, IconWarn, IconCheck } from './Icons'

interface Props {
  onImport: (url: string) => Promise<{ added: number; updated?: number }>
}

/**
 * A single input + button that downloads a pasted Google Sheets / Drive link
 * and imports it: the first import creates the works database; later ones only
 * append works that are new in the sheet (see core/worksSync.ts).
 */
export default function GoogleLinkImport({ onImport }: Props) {
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
      const { added, updated = 0 } = await onImport(trimmed)
      setDone(
        added === 0 && updated === 0
          ? 'No new works in the link — everything in it is already in the database (or the sheet has no rows yet).'
          : `Added ${added} new work${added === 1 ? '' : 's'}, updated ${updated} (Wincode / Eoffice / Estimate amount only) — everything else was left as it is.`
      )
      setUrl('')
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="link-import">
      <p className="link-import-title">Import the works database from a Google Sheets / Drive link</p>
      <div className="link-import-row">
        <IconLink />
        <input
          value={url}
          placeholder="Paste a Google Sheets or Drive link…"
          onChange={(e) => {
            setUrl(e.target.value)
            setError(null)
            setDone(null)
          }}
          onKeyDown={(e) => e.key === 'Enter' && run()}
          disabled={loading}
        />
        <button className="primary" onClick={run} disabled={loading || !url.trim()}>
          {loading ? 'Importing…' : 'Import'}
        </button>
      </div>
      <p className="hint">The first import creates the database. Importing again adds works that are new in the sheet and only updates Wincode, Eoffice and Estimate amount on existing ones — to start over, delete the database first.</p>
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

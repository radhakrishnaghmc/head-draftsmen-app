import { useState } from 'react'
import { IconLink, IconWarn, IconCheck } from './Icons'

interface Props {
  onImport: (url: string) => Promise<{ added: number }>
}

/**
 * A single input + button that downloads a pasted Google Sheets / Drive link
 * and replaces the works list with it — the existing columns and rows are
 * dropped entirely in favor of a fresh database built from the link.
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
      const { added } = await onImport(trimmed)
      setDone(`Replaced the works database with ${added} row${added === 1 ? '' : 's'} from the link.`)
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
      <p className="hint">Importing replaces the current works database with the data from the link.</p>
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

import { useEffect, useState } from 'react'
import { api } from '../ipc'
import { IconDoc, IconDownload, IconWarn, IconRefresh, IconSearch } from './Icons'
import type { CementSteelRate } from '@core/cementSteelRates'
import type { ThemeId } from '../theme'

interface Props {
  /** Called with the freshly-fetched list whenever a load succeeds, so the caller can record what's now been seen. */
  onLoaded?: (rates: CementSteelRate[]) => void
  /** Issue Documents tile style, set in Settings → Themes — applied to this tile grid too, so it matches the rest of the app. */
  theme: ThemeId
}

function sanitizeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim()
}

const MONTHS = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december'
]

/**
 * The rate period a circular actually covers, as a single comparable number
 * (year * 12 + monthIndex). The department batch-uploads several months at
 * once (e.g. Dec 2025, May/Apr/Mar 2026 all "Posted" the same day), so the
 * "File Posted" date doesn't reflect the circular's own period — read the
 * month/year straight out of its description instead, and only fall back to
 * "File Posted" for the rare entry (e.g. an old quarterly one) with no plain
 * month name in its description.
 */
function periodKey(r: CementSteelRate): number {
  const monthMatch = /(january|february|march|april|may|june|july|august|september|october|november|december)\D{0,10}(\d{4})/i.exec(
    r.description
  )
  if (monthMatch) return Number(monthMatch[2]) * 12 + MONTHS.indexOf(monthMatch[1].toLowerCase())
  const dm = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(r.datePosted.trim())
  if (dm) return Number(dm[3]) * 12 + (Number(dm[2]) - 1)
  return 0
}

/**
 * Browse the Telangana Public Health department's "Cement & Steel Rates"
 * circulars (Downloads section of publichealth.telangana.gov.in) as tiles,
 * and download any one of them straight from the source site.
 */
export default function CementSteelRatesPage({ onLoaded, theme }: Props) {
  const [rates, setRates] = useState<CementSteelRate[] | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [search, setSearch] = useState('')
  const [downloadingToken, setDownloadingToken] = useState<string | null>(null)
  const [downloadError, setDownloadError] = useState<string | null>(null)

  async function load() {
    setLoading(true)
    setError(null)
    try {
      const fetched = [...(await api.fetchCementSteelRates())].sort((a, b) => periodKey(b) - periodKey(a))
      setRates(fetched)
      onLoaded?.(fetched)
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    void load()
  }, [])

  async function downloadOne(rate: CementSteelRate) {
    setDownloadingToken(rate.token)
    setDownloadError(null)
    try {
      await api.downloadCementSteelRate(rate.token, `${sanitizeFileName(rate.description)}.${rate.ext}`)
    } catch (e) {
      setDownloadError(e instanceof Error ? e.message : String(e))
    } finally {
      setDownloadingToken(null)
    }
  }

  const filtered = (rates ?? []).filter((r) =>
    search.trim() ? r.description.toLowerCase().includes(search.trim().toLowerCase()) : true
  )

  return (
    <>
      <div className="card pdf-workspace">
        <div className="gps-toolbar">
          <button className="pdf-ws-clearbtn" onClick={load} disabled={loading}>
            <IconRefresh /> {loading ? 'Fetching…' : 'Refresh from website'}
          </button>
          {rates && rates.length > 0 && (
            <div className="mb-search">
              <IconSearch />
              <input
                placeholder="Search by month / year…"
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
              {search && (
                <button className="mb-search-clear" title="Clear search" onClick={() => setSearch('')}>
                  Clear
                </button>
              )}
            </div>
          )}
          {rates && (
            <span className="gps-counts">
              <strong>{filtered.length}</strong> of <strong>{rates.length}</strong> circulars
            </span>
          )}
        </div>

        {error && (
          <div className="notice error">
            <IconWarn /> Couldn't load rates from publichealth.telangana.gov.in — {error}
          </div>
        )}
        {downloadError && (
          <div className="notice error">
            <IconWarn /> {downloadError}
          </div>
        )}
        {loading && !rates && <p className="estimate-hint">Fetching the Cement & Steel Rates list from the department website…</p>}

        {rates && rates.length > 0 && (
          <div className={`doc-tile-grid tools-grid${theme === 'flat1' ? ' tools-grid-flat' : ''}`}>
            {filtered.map((r) => (
              <button
                key={r.token}
                className="doc-tile-card tone-sky tool-card"
                onClick={() => downloadOne(r)}
                disabled={downloadingToken !== null}
                title={r.description}
              >
                <span className="tool-card-ic">
                  <IconDoc />
                </span>
                <span className="doc-tile-card-name">{r.description}</span>
                <span className="doc-tile-card-meta">Posted {r.datePosted}</span>
                <span className="tool-card-cta">
                  <IconDownload /> {downloadingToken === r.token ? 'Saving…' : 'Download'}
                </span>
              </button>
            ))}
          </div>
        )}
      </div>
    </>
  )
}

import { useEffect, useState } from 'react'
import { api } from '../ipc'
import { IconDoc, IconDownload, IconWarn, IconRefresh, IconSearch } from './Icons'
import type { CementSteelRate } from '@core/cementSteelRates'
import { cementSteelRatePeriodKey } from '@core/cementSteelRates'

interface Props {
  /** Called with the freshly-fetched list whenever a load succeeds, so the caller can record what's now been seen. */
  onLoaded?: (rates: CementSteelRate[]) => void
}

function sanitizeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim()
}

/**
 * Browse the Telangana Public Health department's "Cement & Steel Rates"
 * circulars (Downloads section of publichealth.telangana.gov.in) as tiles,
 * and download any one of them straight from the source site.
 */
export default function CementSteelRatesPage({ onLoaded }: Props) {
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
      const fetched = [...(await api.fetchCementSteelRates())].sort((a, b) => cementSteelRatePeriodKey(b) - cementSteelRatePeriodKey(a))
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
          <div className="doc-tile-grid tools-grid">
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

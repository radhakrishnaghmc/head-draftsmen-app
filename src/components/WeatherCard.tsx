import { useEffect, useState } from 'react'
import { api } from '../ipc'
import type { WeatherSnapshot } from '@core/weather'
import {
  IconRefresh,
  IconWarn,
  IconSun,
  IconCloud,
  IconCloudSun,
  IconCloudRain,
  IconCloudLightning,
  IconSnowflake,
  IconMist
} from './Icons'

// All offices this app serves (GHMC/CMC/MMC circles) sit within Hyderabad's
// metro area, so one city query covers every office — no per-circle lookup.
const CITY = 'Hyderabad,IN'

const API_KEY = import.meta.env.VITE_OPENWEATHER_API_KEY

// OpenWeatherMap icon code -> one of the app's own line icons + the page-ic
// colour tone it's shown in, so the card matches the same icon language and
// palette as every page header instead of a platform emoji.
const CONDITION_BY_CODE: Record<string, { Icon: typeof IconSun; tone: string }> = {
  '01': { Icon: IconSun, tone: 'amber' },
  '02': { Icon: IconCloudSun, tone: 'amber' },
  '03': { Icon: IconCloud, tone: '' },
  '04': { Icon: IconCloud, tone: '' },
  '09': { Icon: IconCloudRain, tone: 'sky' },
  '10': { Icon: IconCloudRain, tone: 'sky' },
  '11': { Icon: IconCloudLightning, tone: 'rose' },
  '13': { Icon: IconSnowflake, tone: 'sky' },
  '50': { Icon: IconMist, tone: '' }
}

function conditionFor(icon: string) {
  return CONDITION_BY_CODE[icon.slice(0, 2)] ?? { Icon: IconCloud, tone: '' }
}

/** Dashboard tile: current Hyderabad weather from OpenWeatherMap, refetched (and cached in the main process) every dashboard visit. */
export default function WeatherCard() {
  const [weather, setWeather] = useState<WeatherSnapshot | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    let cancelled = false
    if (!API_KEY) {
      setError('Weather API key not configured.')
      setLoading(false)
      return
    }
    setLoading(true)
    setError(null)
    api
      .fetchWeather(API_KEY, CITY)
      .then((data) => {
        if (!cancelled) setWeather(data)
      })
      .catch((e: unknown) => {
        if (!cancelled) setError(e instanceof Error ? e.message : 'Could not load weather.')
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })
    return () => {
      cancelled = true
    }
  }, [])

  const condition = weather ? conditionFor(weather.icon) : null

  return (
    <div className="card ov-card ov-weather">
      <div className="card-head">
        <h3>Weather</h3>
      </div>
      {loading ? (
        <div className="empty">
          <IconRefresh className="ov-mf-tile-spin" />
          <p>Loading…</p>
        </div>
      ) : error ? (
        <div className="empty">
          <IconWarn />
          <p>{error}</p>
        </div>
      ) : weather && condition ? (
        <div className="ov-weather-body">
          <div className={`page-ic ${condition.tone}`.trim()}>
            <condition.Icon />
          </div>
          <div className="ov-weather-main">
            <b className="ov-weather-temp">{weather.tempC}°C</b>
            <span className="ov-weather-desc">{weather.description}</span>
          </div>
          <ul className="ov-weather-meta">
            <li>{weather.city}</li>
            <li>Feels like {weather.feelsLikeC}°C</li>
            <li>Humidity {weather.humidity}%</li>
            <li>Wind {weather.windKph} km/h</li>
          </ul>
        </div>
      ) : null}
    </div>
  )
}

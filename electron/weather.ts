// Fetches current conditions from OpenWeatherMap's free "current weather
// data" endpoint for the Dashboard's Weather card. Runs in the main process
// (not the renderer) because the app's CSP has no connect-src allowance for
// external hosts — every outside HTTP call goes through main, same as
// fetchCementSteelRates and fetchCalendar.

import type { WeatherSnapshot } from '../core/weather'

const ENDPOINT = 'https://api.openweathermap.org/data/2.5/weather'

interface OwmResponse {
  name?: string
  main?: { temp?: number; feels_like?: number; humidity?: number }
  weather?: { description?: string; icon?: string }[]
  wind?: { speed?: number }
}

export async function fetchWeather(apiKey: string, city: string): Promise<WeatherSnapshot> {
  if (!apiKey) throw new Error('OpenWeatherMap API key not configured — set VITE_OPENWEATHER_API_KEY in a local .env file.')
  const url = `${ENDPOINT}?q=${encodeURIComponent(city)}&appid=${encodeURIComponent(apiKey)}&units=metric`
  const res = await fetch(url)
  if (!res.ok) {
    const body = await res.text().catch(() => '')
    let message = body
    try {
      message = (JSON.parse(body) as { message?: string }).message ?? body
    } catch {
      // body wasn't JSON — use it as-is
    }
    throw new Error(`OpenWeatherMap request failed (${res.status}): ${message || res.statusText}`)
  }
  const data = (await res.json()) as OwmResponse
  return {
    city: data.name || city,
    tempC: Math.round(data.main?.temp ?? 0),
    feelsLikeC: Math.round(data.main?.feels_like ?? 0),
    description: data.weather?.[0]?.description ?? '',
    icon: data.weather?.[0]?.icon ?? '',
    humidity: data.main?.humidity ?? 0,
    // OpenWeatherMap wind speed is in m/s (metric units) — km/h for display.
    windKph: Math.round((data.wind?.speed ?? 0) * 3.6),
    fetchedAt: new Date().toISOString()
  }
}

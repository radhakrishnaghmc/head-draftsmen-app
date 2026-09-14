/** One current-conditions reading for the Dashboard's Weather card (OpenWeatherMap). */
export interface WeatherSnapshot {
  city: string
  tempC: number
  feelsLikeC: number
  description: string
  /** OpenWeatherMap icon code (e.g. "01d") — mapped to one of the app's own line icons in WeatherCard, never loaded as a remote image (CSP blocks external img-src). */
  icon: string
  humidity: number
  windKph: number
  fetchedAt: string
}

/// <reference types="vite/client" />

interface ImportMetaEnv {
  // Local-only dev auto-login password (see src/components/LoginPage.tsx);
  // set in a gitignored .env(.local), never present in a real build.
  readonly VITE_DEV_PASSWORD?: string
  // OpenWeatherMap API key for the Dashboard weather card (see
  // src/components/OverviewDashboard.tsx); set in a gitignored .env(.local).
  readonly VITE_OPENWEATHER_API_KEY?: string
}

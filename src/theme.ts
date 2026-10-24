// Issue Documents tile style — a user-facing preference (Settings → Themes),
// not an office-scoped one like TEMPLATE_KEYS: it's about how this machine
// displays the app, not which document content a given office needs, so it
// stays a plain localStorage value (same "survives updates, not a reinstall"
// tier as LAST_SEEN_VERSION_KEY in App.tsx) rather than office-scoped or
// synced state.
export type ThemeId = 'default' | 'aurora' | 'paper' | 'cyan' | 'play'

const THEME_KEY = 'issueDocsTheme'
// 'default' and 'aurora' are pulled out of the app (not offered in Settings,
// not selectable) but kept here — CSS/types/App.tsx class-toggling untouched
// — so they can be brought back later without rebuilding them from scratch.
// Anyone with one of those already stored falls back to 'paper' below.
const VALID: ThemeId[] = ['play', 'paper', 'cyan']

export function getStoredTheme(): ThemeId {
  const stored = localStorage.getItem(THEME_KEY)
  return (VALID as string[]).includes(stored || '') ? (stored as ThemeId) : 'play'
}

export function setStoredTheme(theme: ThemeId): void {
  localStorage.setItem(THEME_KEY, theme)
}

// Full catalog, including the retired 'default'/'aurora' entries — not
// exported for use in Settings (see THEME_OPTIONS below), kept only so
// re-enabling a retired theme is a one-line change.
const ALL_THEME_OPTIONS: { id: ThemeId; label: string; description: string }[] = [
  {
    id: 'play',
    label: 'Play',
    description:
      'Google Play Console look (Material 3) — pale blue-gray canvas, pill-shaped navigation with a light-blue active item, a white rounded content panel, soft tonal cards and solid blue pill buttons.'
  },
  { id: 'default', label: 'Beginning', description: 'Document thumbnail tiles with a frosted-glass look.' },
  { id: 'aurora', label: 'Aurora', description: 'Whole app restyled with a light background, a vivid cyan accent, and soft pastel tones — a modern SaaS dashboard look.' },
  { id: 'paper', label: 'Paper', description: 'Flat white theme — no glass or blur, thin borders instead of shadows, a single soft blue accent, and a two-column Overview Dashboard layout.' },
  {
    id: 'cyan',
    label: 'Cyan',
    description:
      "Whole app restyled with CMC (cmc.telangana.gov.in)'s own blue — same light glass look as Aurora, just recolored to match the corporation's website."
  }
]

export const THEME_OPTIONS = ALL_THEME_OPTIONS.filter((opt) => (VALID as string[]).includes(opt.id))

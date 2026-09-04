// Issue Documents tile style — a user-facing preference (Settings → Themes),
// not an office-scoped one like TEMPLATE_KEYS: it's about how this machine
// displays the app, not which document content a given office needs, so it
// stays a plain localStorage value (same "survives updates, not a reinstall"
// tier as LAST_SEEN_VERSION_KEY in App.tsx) rather than office-scoped or
// synced state.
export type ThemeId = 'default' | 'aurora'

const THEME_KEY = 'issueDocsTheme'
const VALID: ThemeId[] = ['default', 'aurora']

export function getStoredTheme(): ThemeId {
  const stored = localStorage.getItem(THEME_KEY)
  return (VALID as string[]).includes(stored || '') ? (stored as ThemeId) : 'default'
}

export function setStoredTheme(theme: ThemeId): void {
  localStorage.setItem(THEME_KEY, theme)
}

export const THEME_OPTIONS: { id: ThemeId; label: string; description: string }[] = [
  { id: 'default', label: 'Default', description: 'Document thumbnail tiles with a frosted-glass look.' },
  { id: 'aurora', label: 'Aurora', description: 'Whole app restyled with a light lavender background, a vivid purple accent, and soft pastel tones — a modern SaaS dashboard look.' }
]

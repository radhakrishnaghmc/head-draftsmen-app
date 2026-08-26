// Issue Documents tile style — a user-facing preference (Settings → Themes),
// not an office-scoped one like TEMPLATE_KEYS: it's about how this machine
// displays the app, not which document content a given office needs, so it
// stays a plain localStorage value (same "survives updates, not a reinstall"
// tier as LAST_SEEN_VERSION_KEY in App.tsx) rather than office-scoped or
// synced state.
export type ThemeId = 'default' | 'flat1' | 'windows' | 'dark'

const THEME_KEY = 'issueDocsTheme'
const VALID: ThemeId[] = ['default', 'flat1', 'windows', 'dark']

export function getStoredTheme(): ThemeId {
  const stored = localStorage.getItem(THEME_KEY)
  return (VALID as string[]).includes(stored || '') ? (stored as ThemeId) : 'default'
}

export function setStoredTheme(theme: ThemeId): void {
  localStorage.setItem(THEME_KEY, theme)
}

export const THEME_OPTIONS: { id: ThemeId; label: string; description: string }[] = [
  { id: 'default', label: 'Default', description: 'Document thumbnail tiles with a frosted-glass look.' },
  { id: 'flat1', label: 'Colured', description: 'Flat, solid-color tiles with a large icon.' },
  { id: 'windows', label: 'Windows theme', description: 'Flat, sharp-cornered surfaces styled after Windows 11 Settings.' },
  { id: 'dark', label: 'Dark mode', description: 'Whole app restyled like iOS Dark Mode — true-black surfaces, frosted sidebar, system-blue buttons.' }
]

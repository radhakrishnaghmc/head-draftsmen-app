import { useEffect, useRef, useState } from 'react'
import { IconUser, IconLogout, IconRefresh, IconCheck } from './Icons'

interface Props {
  /** The signed-in user's Login ID, shown in the dropdown and on the Log out button. */
  username: string
  onLogout: () => void
  /** Pushes the workspace to the cloud now; resolves '' when it was saved online, otherwise the reason it failed. */
  onSync: () => Promise<string>
  /** Edits not yet in the cloud, in bytes; the app syncs by itself once this reaches autoSyncBytes. */
  pendingBytes: number
  autoSyncBytes: number
}

/** Fixed top-right profile badge — click to open a small menu with Log out. */
export default function ProfileMenu({ username, onLogout, onSync, pendingBytes, autoSyncBytes }: Props) {
  const [open, setOpen] = useState(false)
  const [sync, setSync] = useState<'idle' | 'busy' | 'done' | 'failed'>('idle')
  const [syncError, setSyncError] = useState('')
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!open) return
    function onClickOutside(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown', onClickOutside)
    return () => document.removeEventListener('mousedown', onClickOutside)
  }, [open])

  const fmtKb = (b: number) => (b < 1024 ? '<1 KB' : `${(b / 1024).toFixed(b < 10240 ? 1 : 0)} KB`)
  const pendingText =
    pendingBytes > 0
      ? `${fmtKb(pendingBytes)} not synced (auto-syncs at ${fmtKb(autoSyncBytes)})`
      : 'Everything is synced'
  const syncLabel =
    sync === 'busy'
      ? 'Syncing…'
      : sync === 'done'
        ? 'Synced to online'
        : sync === 'failed'
          ? 'Sync failed'
          : 'Sync to online'

  return (
    <div className="profile-menu" ref={ref}>
      <button
        className={`profile-badge profile-sync profile-sync-${sync}`}
        disabled={sync === 'busy'}
        title={sync === 'idle' ? `${syncLabel} — ${pendingText}` : syncLabel}
        aria-label={syncLabel}
        onClick={() => {
          setOpen(false)
          setSync('busy')
          void onSync().then((err) => {
            setSyncError(err)
            setSync(err ? 'failed' : 'done')
            setTimeout(() => setSync('idle'), err ? 12000 : 3000)
          })
        }}
      >
        {sync === 'done' ? <IconCheck /> : <IconRefresh />}
      </button>
      {pendingBytes > 0 && sync === 'idle' && <span className="profile-sync-pending">{fmtKb(pendingBytes)}</span>}
      <button className="profile-badge" onClick={() => setOpen((o) => !o)} title={username || 'Account'}>
        <IconUser />
      </button>
      {open && (
        <div className="profile-dropdown">
          {username && <div className="profile-dropdown-user">{username}</div>}
          <button
            className="profile-dropdown-item"
            onClick={() => {
              setOpen(false)
              onLogout()
            }}
          >
            <IconLogout /> Log out{username && ` (${username})`}
          </button>
        </div>
      )}
      {sync === 'failed' && syncError && !open && (
        <div className="profile-dropdown profile-sync-error">{syncError.slice(0, 160)}</div>
      )}
    </div>
  )
}

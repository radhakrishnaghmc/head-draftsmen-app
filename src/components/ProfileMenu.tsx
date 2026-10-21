import { useEffect, useRef, useState } from 'react'
import { IconUser, IconLogout, IconRefresh } from './Icons'

interface Props {
  /** The signed-in user's Login ID, shown in the dropdown and on the Log out button. */
  username: string
  onLogout: () => void
  /** Pushes the workspace to the cloud now; resolves true when it was saved online. */
  onSync: () => Promise<boolean>
}

/** Fixed top-right profile badge — click to open a small menu with Log out. */
export default function ProfileMenu({ username, onLogout, onSync }: Props) {
  const [open, setOpen] = useState(false)
  const [sync, setSync] = useState<'idle' | 'busy' | 'done' | 'failed'>('idle')
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!open) return
    function onClickOutside(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown', onClickOutside)
    return () => document.removeEventListener('mousedown', onClickOutside)
  }, [open])

  return (
    <div className="profile-menu" ref={ref}>
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
          <button
            className="profile-dropdown-item profile-dropdown-sync"
            disabled={sync === 'busy'}
            onClick={() => {
              setSync('busy')
              void onSync().then((ok) => {
                setSync(ok ? 'done' : 'failed')
                setTimeout(() => setSync('idle'), 3000)
              })
            }}
          >
            <IconRefresh /> {sync === 'busy' ? 'Syncing…' : sync === 'done' ? 'Synced to online' : sync === 'failed' ? 'Sync failed — check connection' : 'Sync across devices'}
          </button>
        </div>
      )}
    </div>
  )
}

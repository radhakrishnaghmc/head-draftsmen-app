import { useEffect, useRef, useState } from 'react'
import { IconUser, IconLogout } from './Icons'

interface Props {
  /** The signed-in user's Login ID, shown in the dropdown and on the Log out button. */
  username: string
  onLogout: () => void
}

/** Fixed top-right profile badge — click to open a small menu with Log out. */
export default function ProfileMenu({ username, onLogout }: Props) {
  const [open, setOpen] = useState(false)
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
        </div>
      )}
    </div>
  )
}

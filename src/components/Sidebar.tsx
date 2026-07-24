import { useEffect, useState } from 'react'
import { api } from '../ipc'
import {
  IconTable,
  IconDoc,
  IconPrint,
  IconLogo,
  IconWarn,
  IconCheck,
  IconSearch,
  IconCalendar,
  IconDownload,
  IconChecklist
} from './Icons'

export type TabKey =
  | 'dashboard'
  | 'data'
  | 'createDoc'
  | 'printDoc'
  | 'boq'
  | 'scheduleA'
  | 'deviation'
  | 'techSanction'
  | 'search'
  | 'todo'

interface Props {
  active: TabKey
  onSelect: (t: TabKey) => void
  tableCount: number
  unresolved: number
  createdDocCount: number
  /** From the Works List's first row — shown under the app name. */
  zone?: string
  circle?: string
}

interface Item {
  key: TabKey
  label: string
  icon: JSX.Element
  badge?: string | number
  warn?: boolean
  ok?: boolean
  tone?: 'green' | 'sky' | 'rose' | 'amber' | 'teal'
}

export default function Sidebar(props: Props) {
  const [version, setVersion] = useState<string | null>(null)
  useEffect(() => {
    api.getAppVersion().then(setVersion)
  }, [])

  const items: Item[] = [
    {
      key: 'dashboard',
      label: 'Calendar',
      icon: <IconCalendar />
    },
    {
      key: 'data',
      label: 'Works List',
      icon: <IconTable />,
      badge: props.tableCount || undefined,
      warn: props.unresolved > 0,
      tone: 'green'
    },
    {
      key: 'createDoc',
      label: 'Create New Document',
      icon: <IconDoc />,
      tone: 'rose'
    },
    {
      key: 'printDoc',
      label: 'Issue Document',
      icon: <IconPrint />,
      badge: props.createdDocCount || undefined,
      tone: 'sky'
    },
    {
      key: 'boq',
      label: 'Download BOQ',
      icon: <IconDownload />,
      tone: 'amber'
    },
    {
      key: 'scheduleA',
      label: 'Download Schedule A',
      icon: <IconDownload />,
      tone: 'teal'
    },
    {
      key: 'deviation',
      label: 'Get Deviation Statement',
      icon: <IconDownload />
    },
    {
      key: 'techSanction',
      label: 'Give Technical Sanction',
      icon: <IconCheck />
    }
  ]

  return (
    <aside className="sidebar">
      <div className="side-brand">
        <IconLogo className="logo" />
        <div>
          <h1>Head Draftsman</h1>
          {(props.zone || props.circle) && (
            <p>{[props.circle, props.zone].filter(Boolean).join(' · ')}</p>
          )}
        </div>
      </div>

      <nav className="side-nav">
        {items.map((it, i) => (
          <button
            key={it.key}
            className={`nav-item ${props.active === it.key ? 'on' : ''} ${it.tone ? `tone-${it.tone}` : ''}`}
            onClick={() => props.onSelect(it.key)}
          >
            <span className="nav-step">{i + 1}</span>
            <span className="nav-ic">{it.icon}</span>
            <span className="nav-text">
              <span className="nav-label">{it.label}</span>
            </span>
            {it.warn ? (
              <span className="nav-badge warn">
                <IconWarn />
              </span>
            ) : it.ok ? (
              <span className="nav-badge ok">
                <IconCheck />
              </span>
            ) : it.badge != null ? (
              <span className="nav-badge">{it.badge}</span>
            ) : null}
          </button>
        ))}

        <div className="nav-sep" />

        <button
          className={`nav-item ${props.active === 'todo' ? 'on' : ''}`}
          onClick={() => props.onSelect('todo')}
        >
          <span className="nav-step tool">
            <IconChecklist />
          </span>
          <span className="nav-ic">
            <IconChecklist />
          </span>
          <span className="nav-text">
            <span className="nav-label">To Do List</span>
          </span>
        </button>

        <button
          className={`nav-item tone-green ${props.active === 'search' ? 'on' : ''}`}
          onClick={() => props.onSelect('search')}
        >
          <span className="nav-step tool">
            <IconSearch />
          </span>
          <span className="nav-ic">
            <IconSearch />
          </span>
          <span className="nav-text">
            <span className="nav-label">Search Tender</span>
          </span>
        </button>
      </nav>

      <div className="side-foot">
        <span className="side-credits">
          App developed by Radhakrishna, HD{version ? ` · v${version}` : ''}
        </span>
      </div>
    </aside>
  )
}

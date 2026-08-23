import { useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import { IconPlus, IconTrash, IconCheck, IconRefresh, IconEye, IconChevronRight, IconSearch } from './Icons'
import type { MBScrutinyItem } from '@core/types'
import { closeOnBackdropMouseDown } from '../overlayClose'

function nextSerialNo(items: MBScrutinyItem[]): number {
  return items.reduce((max, it) => Math.max(max, it.serialNo || 0), 0) + 1
}

interface Props {
  items: MBScrutinyItem[]
  onChange: (items: MBScrutinyItem[]) => void
}

function todayISO(): string {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

function formatDDMMYYYY(iso: string): string {
  if (!iso) return ''
  const [y, m, d] = iso.split('-')
  return `${d}.${m}.${y}`
}

function nextId(): string {
  return `mb-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

/**
 * MB Scrutiny list: a log of Measurement Books received for verification.
 * Each entry records the MB no., agency, and the date it was received;
 * marking it done stamps a scrutiny-completed date but keeps the entry in
 * the list permanently — this is a record, not a rolling task list.
 */
export default function MbScrutinyList({ items, onChange }: Props) {
  const [mbNo, setMbNo] = useState('')
  const [agencyName, setAgencyName] = useState('')
  const [receivedDate, setReceivedDate] = useState(todayISO())
  const [targetDate, setTargetDate] = useState('')
  const [remarkDrafts, setRemarkDrafts] = useState<Record<string, string>>({})
  const [pendingDelete, setPendingDelete] = useState<MBScrutinyItem | null>(null)
  const [search, setSearch] = useState('')
  // Behaves like a tab group: at most one list is shown at a time, so the
  // completed list never stacks under the pending one. Clicking the active
  // tab again collapses it (shows neither list).
  const [activeTab, setActiveTab] = useState<'pending' | 'done' | null>('pending')

  function toggleTab(tab: 'pending' | 'done') {
    setActiveTab((cur) => (cur === tab ? null : tab))
  }

  // Checking an item's box only flips `done` — its serialNo (the register
  // token number) was assigned once at creation and is never reassigned, so
  // moving between these two sections never renumbers anything.
  const query = search.trim().toLowerCase()
  const matchesSearch = (it: MBScrutinyItem) =>
    !query || it.mbNo.toLowerCase().includes(query) || it.agencyName.toLowerCase().includes(query)

  const pendingItems = useMemo(
    () =>
      items
        .filter((it) => !it.done && matchesSearch(it))
        .sort((a, b) => (a.receivedDate < b.receivedDate ? -1 : a.receivedDate > b.receivedDate ? 1 : 0)),
    [items, query]
  )
  const doneItems = useMemo(
    () =>
      items
        .filter((it) => it.done && matchesSearch(it))
        .sort((a, b) => (b.completedDate ?? '').localeCompare(a.completedDate ?? '')),
    [items, query]
  )

  function addItem() {
    const trimmedMbNo = mbNo.trim()
    const trimmedAgency = agencyName.trim()
    if (!trimmedMbNo || !trimmedAgency || !receivedDate) return
    onChange([
      ...items,
      {
        id: nextId(),
        serialNo: nextSerialNo(items),
        mbNo: trimmedMbNo,
        agencyName: trimmedAgency,
        receivedDate,
        targetDate: targetDate || undefined,
        done: false
      }
    ])
    setMbNo('')
    setAgencyName('')
    setReceivedDate(todayISO())
    setTargetDate('')
  }

  function update(id: string, patch: Partial<MBScrutinyItem>) {
    onChange(items.map((it) => (it.id === id ? { ...it, ...patch } : it)))
  }

  function toggleDone(id: string) {
    onChange(
      items.map((it) =>
        it.id === id
          ? it.done
            ? { ...it, done: false, completedDate: undefined }
            : { ...it, done: true, completedDate: todayISO() }
          : it
      )
    )
  }

  function confirmDelete() {
    if (!pendingDelete) return
    onChange(items.filter((it) => it.id !== pendingDelete.id))
    setPendingDelete(null)
  }

  function addRemark(id: string) {
    const draft = (remarkDrafts[id] ?? '').trim()
    if (!draft) return
    const item = items.find((it) => it.id === id)
    update(id, { remarks: [...(item?.remarks ?? []), { text: draft, done: false }] })
    setRemarkDrafts((d) => ({ ...d, [id]: '' }))
  }

  function toggleRemarkDone(id: string, index: number) {
    const item = items.find((it) => it.id === id)
    update(id, {
      remarks: (item?.remarks ?? []).map((r, i) => (i === index ? { ...r, done: !r.done } : r))
    })
  }

  function removeRemark(id: string, index: number) {
    const item = items.find((it) => it.id === id)
    update(id, { remarks: (item?.remarks ?? []).filter((_, i) => i !== index) })
  }

  function renderRow(it: MBScrutinyItem) {
    return (
      <li key={it.id} className={`todo-item mb-item ${it.done ? 'done' : ''}`}>
        <input
          type="checkbox"
          className="todo-checkbox"
          checked={it.done}
          title={it.done ? 'Reopen' : 'Mark scrutiny done'}
          onChange={() => toggleDone(it.id)}
        />
        <span className="mb-serial" title="Register S.No. — fixed for this MB">
          #{it.serialNo}
        </span>
        <div className="todo-body">
          <span className="todo-text">
            MB No. {it.mbNo} · {it.agencyName}
          </span>
          <span className="todo-dates">
            Received {formatDDMMYYYY(it.receivedDate)}
            {it.targetDate && ` · Target ${formatDDMMYYYY(it.targetDate)}`}
            {it.done && it.completedDate && ` · Scrutiny completed ${formatDDMMYYYY(it.completedDate)}`}
          </span>
          <div className="mb-remarks">
            {it.remarks && it.remarks.length > 0 && (
              <ul className="mb-remarks-points">
                {it.remarks.map((r, i) => (
                  <li key={i} className={r.done ? 'done' : ''}>
                    <input
                      type="checkbox"
                      checked={r.done}
                      title={r.done ? 'Mark unattended' : 'Mark attended'}
                      onChange={() => toggleRemarkDone(it.id, i)}
                    />
                    <span>{r.text}</span>
                    <button className="mb-remark-del" title="Remove point" onClick={() => removeRemark(it.id, i)}>
                      <IconTrash />
                    </button>
                  </li>
                ))}
              </ul>
            )}
            <div className="mb-remarks-add">
              <input
                value={remarkDrafts[it.id] ?? ''}
                onChange={(e) => setRemarkDrafts((d) => ({ ...d, [it.id]: e.target.value }))}
                onKeyDown={(e) => e.key === 'Enter' && addRemark(it.id)}
                placeholder="Add a remark / objection point…"
              />
              <button className="ghost" title="Add point" onClick={() => addRemark(it.id)}>
                <IconPlus />
              </button>
            </div>
          </div>
        </div>
        {!it.done ? (
          <label className="mb-row-date">
            <span>Target</span>
            <input
              type="date"
              className="todo-target-input"
              value={it.targetDate ?? ''}
              onChange={(e) => update(it.id, { targetDate: e.target.value })}
              title="Target completion date"
            />
          </label>
        ) : (
          <label className="mb-row-date">
            <span>Completed</span>
            <input
              type="date"
              className="todo-target-input"
              value={it.completedDate ?? ''}
              onChange={(e) => update(it.id, { completedDate: e.target.value })}
              title="Scrutiny completed date"
            />
          </label>
        )}
        <button className="danger-ghost" title="Delete" onClick={() => setPendingDelete(it)}>
          <IconTrash />
        </button>
      </li>
    )
  }

  return (
    <div className="card mb-card">
      <div className="mb-add">
        <label className="mb-field mb-field-grow">
          <span>MB No.</span>
          <input value={mbNo} onChange={(e) => setMbNo(e.target.value)} placeholder="e.g. 120/25-26" />
        </label>
        <label className="mb-field mb-field-grow">
          <span>Agency name</span>
          <input value={agencyName} onChange={(e) => setAgencyName(e.target.value)} placeholder="Agency name" />
        </label>
        <label className="mb-field">
          <span>Received</span>
          <input
            type="date"
            value={receivedDate}
            onChange={(e) => setReceivedDate(e.target.value)}
            title="MB received date"
          />
        </label>
        <label className="mb-field">
          <span>Target</span>
          <input
            type="date"
            value={targetDate}
            onChange={(e) => setTargetDate(e.target.value)}
            title="Target completion date"
          />
        </label>
        <button className="primary" onClick={addItem}>
          <IconPlus /> Add
        </button>
      </div>

      {items.length === 0 ? (
        <div className="empty">
          <IconEye />
          <p>No MBs logged yet — add one above.</p>
        </div>
      ) : (
        <>
          <div className="mb-search">
            <IconSearch />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search by MB number or agency name…"
            />
            {search && (
              <button className="mb-search-clear" title="Clear search" onClick={() => setSearch('')}>
                Clear
              </button>
            )}
          </div>

          <div className="mb-tabs">
            <button
              className={`mb-tab tone-pending ${activeTab === 'pending' ? 'open' : ''}`}
              onClick={() => toggleTab('pending')}
            >
              <span className="mb-section-ic">
                <IconRefresh />
              </span>
              <span className="mb-tab-label">Pending scrutiny</span>
              <span className="mb-section-count">{pendingItems.length}</span>
              <IconChevronRight className="mb-tab-chevron" />
            </button>
            <button
              className={`mb-tab tone-done ${activeTab === 'done' ? 'open' : ''}`}
              onClick={() => toggleTab('done')}
            >
              <span className="mb-section-ic">
                <IconCheck />
              </span>
              <span className="mb-tab-label">Scrutiny completed</span>
              <span className="mb-section-count">{doneItems.length}</span>
              <IconChevronRight className="mb-tab-chevron" />
            </button>
          </div>

          {activeTab === 'pending' &&
            (pendingItems.length === 0 ? (
              <p className="mb-section-empty">
                {query
                  ? `No pending MBs match "${search.trim()}".`
                  : 'Nothing pending — every MB on file has been scrutinised.'}
              </p>
            ) : (
              <ul className="todo-list mb-section">{pendingItems.map(renderRow)}</ul>
            ))}

          {activeTab === 'done' &&
            (doneItems.length === 0 ? (
              <p className="mb-section-empty">
                {query ? `No completed MBs match "${search.trim()}".` : 'Nothing marked done yet.'}
              </p>
            ) : (
              <ul className="todo-list mb-section">{doneItems.map(renderRow)}</ul>
            ))}
        </>
      )}

      {pendingDelete &&
        createPortal(
          <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => setPendingDelete(null))}>
            <div className="confirm-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
              <div className="confirm-ic">
                <IconTrash />
              </div>
              <h3>Delete this MB record?</h3>
              <p className="confirm-warn">
                You're about to permanently remove MB No. <strong>{pendingDelete.mbNo}</strong> (
                {pendingDelete.agencyName}).
              </p>
              <p className="confirm-hint">Once deleted, this cannot be recovered.</p>
              <div className="confirm-actions">
                <button className="ghost" onClick={() => setPendingDelete(null)}>
                  Cancel
                </button>
                <button className="danger" onClick={confirmDelete}>
                  Delete
                </button>
              </div>
            </div>
          </div>,
          document.body
        )}
    </div>
  )
}

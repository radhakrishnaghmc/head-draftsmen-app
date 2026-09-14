import { useMemo, useState } from 'react'
import { IconPlus, IconTrash, IconCheck, IconEdit, IconRestore, IconRefresh, IconChevronRight, IconWarn } from './Icons'
import type { TodoItem } from '@core/types'

interface Props {
  todos: TodoItem[]
  onChange: (todos: TodoItem[]) => void
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
  return `todo-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
}

/**
 * To Do List: add tasks with a target completion date (created date defaults
 * to today). Ticking a task keeps it visible for the rest of that day, then
 * it rolls off the list once the calendar date moves on — incomplete tasks
 * stay until they're done, regardless of date.
 */
export default function TodoList({ todos, onChange }: Props) {
  const today = todayISO()
  const [text, setText] = useState('')
  const [targetDate, setTargetDate] = useState('')
  // The task currently being edited inline, and its working text. null = none.
  const [editingId, setEditingId] = useState<string | null>(null)
  const [editText, setEditText] = useState('')
  // Behaves like a tab group: at most one section is shown at a time, mirroring
  // the MB Scrutiny register's Pending/Done/Bin layout.
  const [activeTab, setActiveTab] = useState<'pending' | 'done' | 'bin' | null>('pending')

  function toggleTab(tab: 'pending' | 'done' | 'bin') {
    setActiveTab((cur) => (cur === tab ? null : tab))
  }

  const pendingItems = useMemo(() => todos.filter((t) => !t.deletedAt && !t.done), [todos])
  const doneTodayItems = useMemo(
    () => todos.filter((t) => !t.deletedAt && t.done && t.completedDate === today),
    [todos, today]
  )
  const binned = useMemo(
    () =>
      todos
        .filter((t) => t.deletedAt)
        .sort((a, b) => (b.deletedAt ?? '').localeCompare(a.deletedAt ?? '')),
    [todos]
  )

  function addItem() {
    const trimmed = text.trim()
    if (!trimmed) return
    onChange([...todos, { id: nextId(), text: trimmed, createdDate: today, targetDate, done: false }])
    setText('')
    setTargetDate('')
  }

  function toggleDone(id: string) {
    onChange(
      todos.map((t) =>
        t.id === id
          ? t.done
            ? { ...t, done: false, completedDate: undefined }
            : { ...t, done: true, completedDate: today }
          : t
      )
    )
  }

  function setItemTarget(id: string, value: string) {
    onChange(todos.map((t) => (t.id === id ? { ...t, targetDate: value } : t)))
  }

  function startEdit(item: TodoItem) {
    setEditingId(item.id)
    setEditText(item.text)
  }

  // Save the edited text (unless blank — an empty task makes no sense, so a
  // blank edit just cancels), then leave edit mode.
  function commitEdit() {
    const trimmed = editText.trim()
    if (editingId && trimmed) {
      onChange(todos.map((t) => (t.id === editingId ? { ...t, text: trimmed } : t)))
    }
    setEditingId(null)
    setEditText('')
  }

  function cancelEdit() {
    setEditingId(null)
    setEditText('')
  }

  function moveToBin(id: string) {
    onChange(todos.map((t) => (t.id === id ? { ...t, deletedAt: new Date().toISOString() } : t)))
  }

  function restoreItem(id: string) {
    onChange(todos.map((t) => (t.id === id ? { ...t, deletedAt: undefined } : t)))
  }

  function renderRow(t: TodoItem) {
    const overdue = !t.done && !!t.targetDate && t.targetDate < today
    return (
      <li key={t.id} className={`todo-item mb-item ${t.done ? 'done' : ''}`}>
        <input
          type="checkbox"
          className="todo-checkbox"
          checked={t.done}
          title={t.done ? 'Reopen' : 'Mark done'}
          onChange={() => toggleDone(t.id)}
        />
        <div className="todo-body">
          {editingId === t.id ? (
            <input
              className="todo-edit-input"
              value={editText}
              autoFocus
              onChange={(e) => setEditText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') commitEdit()
                else if (e.key === 'Escape') cancelEdit()
              }}
              onBlur={commitEdit}
            />
          ) : (
            <span className="todo-text">{t.text}</span>
          )}
          <span className="todo-dates">
            Added {formatDDMMYYYY(t.createdDate)}
            {t.targetDate && ` · Target ${formatDDMMYYYY(t.targetDate)}`}
            {t.done && t.completedDate && ` · Completed ${formatDDMMYYYY(t.completedDate)}`}
            {overdue && (
              <span className="badge err">
                <IconWarn /> Overdue
              </span>
            )}
          </span>
        </div>
        {!t.done && (
          <label className="mb-row-date">
            <span>Target</span>
            <input
              type="date"
              className="todo-target-input"
              value={t.targetDate}
              onChange={(e) => setItemTarget(t.id, e.target.value)}
              title="Target completion date"
            />
          </label>
        )}
        {editingId === t.id ? (
          <button className="ghost" title="Save" onClick={commitEdit}>
            <IconCheck />
          </button>
        ) : (
          <button className="ghost" title="Edit task" onClick={() => startEdit(t)}>
            <IconEdit />
          </button>
        )}
        <button className="danger-ghost" title="Move to Bin" onClick={() => moveToBin(t.id)}>
          <IconTrash />
          Bin
        </button>
      </li>
    )
  }

  return (
    <div className="card mb-card">
      <div className="mb-add">
        <label className="mb-field mb-field-grow">
          <span>Task</span>
          <input
            value={text}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && addItem()}
            placeholder="Add a task…"
          />
        </label>
        <label className="mb-field">
          <span>Target date</span>
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

      {todos.length === 0 ? (
        <div className="empty">
          <IconCheck />
          <p>Nothing to do — add a task above.</p>
        </div>
      ) : (
        <>
          <div className="mb-tabs">
            <button
              className={`mb-tab tone-pending ${activeTab === 'pending' ? 'open' : ''}`}
              onClick={() => toggleTab('pending')}
            >
              <span className="mb-section-ic">
                <IconRefresh />
              </span>
              <span className="mb-tab-label">Pending</span>
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
              <span className="mb-tab-label">Done today</span>
              <span className="mb-section-count">{doneTodayItems.length}</span>
              <IconChevronRight className="mb-tab-chevron" />
            </button>
            <button
              className={`mb-tab tone-bin ${activeTab === 'bin' ? 'open' : ''}`}
              onClick={() => toggleTab('bin')}
            >
              <span className="mb-section-ic">
                <IconTrash />
              </span>
              <span className="mb-tab-label">Bin</span>
              <span className="mb-section-count">{binned.length}</span>
              <IconChevronRight className="mb-tab-chevron" />
            </button>
          </div>

          {activeTab === 'pending' &&
            (pendingItems.length === 0 ? (
              <p className="mb-section-empty">Nothing pending — every task is done.</p>
            ) : (
              <ul className="todo-list mb-section">{pendingItems.map(renderRow)}</ul>
            ))}

          {activeTab === 'done' &&
            (doneTodayItems.length === 0 ? (
              <p className="mb-section-empty">Nothing marked done today yet.</p>
            ) : (
              <ul className="todo-list mb-section">{doneTodayItems.map(renderRow)}</ul>
            ))}

          {activeTab === 'bin' &&
            (binned.length === 0 ? (
              <p className="mb-section-empty">The Bin is empty.</p>
            ) : (
              <ul className="todo-list mb-section">
                {binned.map((t) => (
                  <li key={t.id} className="todo-item mb-item mb-item-binned">
                    <div className="todo-body">
                      <span className="todo-text">{t.text}</span>
                      <span className="todo-dates">
                        Added {formatDDMMYYYY(t.createdDate)}
                        {t.deletedAt && ` · Moved to Bin ${formatDDMMYYYY(t.deletedAt.slice(0, 10))}`}
                      </span>
                    </div>
                    <button className="ghost" title="Restore" onClick={() => restoreItem(t.id)}>
                      <IconRestore />
                      Restore
                    </button>
                  </li>
                ))}
              </ul>
            ))}
        </>
      )}
    </div>
  )
}

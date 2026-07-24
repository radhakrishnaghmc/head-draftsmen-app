import { useMemo, useState } from 'react'
import { IconPlus, IconTrash, IconCheck } from './Icons'
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

  const visible = useMemo(
    () => todos.filter((t) => !t.done || t.completedDate === today),
    [todos, today]
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

  function removeItem(id: string) {
    onChange(todos.filter((t) => t.id !== id))
  }

  return (
    <div className="card">
      <div className="todo-add">
        <input
          value={text}
          onChange={(e) => setText(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && addItem()}
          placeholder="Add a task…"
        />
        <input
          type="date"
          value={targetDate}
          onChange={(e) => setTargetDate(e.target.value)}
          title="Target completion date"
        />
        <button className="primary" onClick={addItem}>
          <IconPlus /> Add
        </button>
      </div>

      {visible.length === 0 ? (
        <div className="empty">
          <IconCheck />
          <p>Nothing to do — add a task above.</p>
        </div>
      ) : (
        <ul className="todo-list">
          {visible.map((t) => (
            <li key={t.id} className={`todo-item ${t.done ? 'done' : ''}`}>
              <input
                type="checkbox"
                className="todo-checkbox"
                checked={t.done}
                onChange={() => toggleDone(t.id)}
              />
              <div className="todo-body">
                <span className="todo-text">{t.text}</span>
                <span className="todo-dates">
                  Added {formatDDMMYYYY(t.createdDate)}
                  {t.targetDate && ` · Target ${formatDDMMYYYY(t.targetDate)}`}
                </span>
              </div>
              <input
                type="date"
                className="todo-target-input"
                value={t.targetDate}
                onChange={(e) => setItemTarget(t.id, e.target.value)}
                title="Target completion date"
              />
              <button className="danger-ghost" title="Delete" onClick={() => removeItem(t.id)}>
                <IconTrash />
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}

import { describe, it, expect } from 'vitest'
import { encryptBackup, decryptBackup, BackupError, mergeActivityLogs, describeStateChange, type BackupPayload } from '../core/backup'
import type { PersistedState } from '../core/types'

const state = { version: 1, tables: [], resolution: {}, todos: [{ id: 't1' }] } as unknown as PersistedState
const payload: BackupPayload = { appVersion: '1.0.0', createdAt: 1, loginId: 'Alice', state, activityLog: [] }

describe('backup encryption', () => {
  it('round-trips under the same login', () => {
    const file = encryptBackup(payload, 'Alice', 'pw')
    expect(decryptBackup(file, ' alice ', 'pw').state).toEqual(state)
  })
  it('does not leak plaintext', () => {
    const secret = { ...payload, state: { ...state, lastGoogleLink: 'SECRET-MARKER-https://example.test' } }
    expect(encryptBackup(secret, 'Alice', 'pw').toString('utf8')).not.toContain('SECRET-MARKER')
  })
  it('rejects another login or password', () => {
    const file = encryptBackup(payload, 'Alice', 'pw')
    expect(() => decryptBackup(file, 'Bob', 'pw')).toThrow(BackupError)
    expect(() => decryptBackup(file, 'Alice', 'other')).toThrow(BackupError)
  })
  it('rejects non-backup and tampered files', () => {
    expect(() => decryptBackup(Buffer.from('hello'), 'Alice', 'pw')).toThrow(/not a Head Draftsmen/)
    const env = JSON.parse(encryptBackup(payload, 'Alice', 'pw').toString())
    env.data = Buffer.from('xx').toString('base64')
    expect(() => decryptBackup(Buffer.from(JSON.stringify(env)), 'Alice', 'pw')).toThrow(BackupError)
  })
})

describe('activity helpers', () => {
  it('merges logs by id, oldest first', () => {
    const a = [{ id: '2', ts: 2, action: 'x', detail: '' }]
    const b = [{ id: '1', ts: 1, action: 'x', detail: '' }, { id: '2', ts: 2, action: 'x', detail: '' }]
    expect(mergeActivityLogs(a, b).map((e) => e.id)).toEqual(['1', '2'])
  })
  it('describes count changes and edits', () => {
    const first = describeStateChange({}, state)
    const next = { ...state, todos: [{ id: 't1' }, { id: 't2' }] } as unknown as PersistedState
    expect(describeStateChange(first.sigs, next).detail).toBe('To Do 1→2')
    expect(describeStateChange(first.sigs, state).detail).toBe('')
  })
})

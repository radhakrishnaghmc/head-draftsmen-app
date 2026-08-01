// Cross-device sync + "max concurrent logins" enforcement, backed by
// Firestore. Runs entirely in the Electron main process (the renderer never
// talks to Firebase directly). Every call here is best-effort: if the cloud
// is unreachable, the app must keep working exactly as it did before this
// feature existed — offline-first, local state.json is always the fallback.
import { initializeApp, type FirebaseApp } from 'firebase/app'
import { getAuth, signInAnonymously, type Auth } from 'firebase/auth'
import {
  getFirestore,
  doc,
  getDoc,
  getDocs,
  setDoc,
  deleteDoc,
  runTransaction,
  collection,
  onSnapshot,
  serverTimestamp,
  type Firestore,
  type Unsubscribe
} from 'firebase/firestore'
import { canClaimSlot, claimSlot, releaseSlot, touchSlot, type SessionSlot } from '../core/sessionSlots'
import type { PersistedState, CreatedDocument } from '../core/types'

const firebaseConfig = {
  apiKey: 'AIzaSyBBoysv9uaHzJmeptn97Ms73AOGauEn0DU',
  authDomain: 'head-draftsmen-app.firebaseapp.com',
  projectId: 'head-draftsmen-app',
  storageBucket: 'head-draftsmen-app.firebasestorage.app',
  messagingSenderId: '71347636667',
  appId: '1:71347636667:web:cdfb6b67902ab8e8c0dc9e'
}

let app: FirebaseApp | null = null
let auth: Auth | null = null
let db: Firestore | null = null
let authReady: Promise<void> | null = null

function ensureAuth(): Promise<void> {
  if (!app) {
    app = initializeApp(firebaseConfig)
    auth = getAuth(app)
    db = getFirestore(app)
  }
  if (!authReady) {
    authReady = signInAnonymously(auth!)
      .then(() => undefined)
      .catch((e) => {
        authReady = null
        throw e
      })
  }
  return authReady
}

function normalizeId(loginId: string): string {
  return loginId.trim().toLowerCase().replace(/[^a-z0-9_-]/g, '_') || 'unknown'
}

export interface ClaimResult {
  ok: boolean
  /** True only when the cloud was reachable and genuinely already has 2 live sessions. */
  maxSessions?: boolean
}

interface ActiveSession {
  loginId: string
  sessionId: string
  knownDocIds: Set<string>
  unsubscribes: Unsubscribe[]
  heartbeatTimer: ReturnType<typeof setInterval> | null
}

let active: ActiveSession | null = null

/**
 * Call after a successful password check. Claims one of the 2 session slots
 * for this login ID, pulls whatever's already synced for it, and starts a
 * live listener for changes made from the other device.
 *
 * If the cloud is unreachable, the login proceeds anyway (offline-only, no
 * sync, no session cap) rather than locking the user out over a network hiccup.
 */
export async function startSession(
  loginId: string,
  deviceLabel: string,
  onRemoteUpdate: (partial: Partial<PersistedState>) => void
): Promise<{ claim: ClaimResult; remoteState: PersistedState | null }> {
  const id = normalizeId(loginId)
  try {
    await ensureAuth()
    const ref = doc(db!, 'sessions', id)
    const sessionId = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
    const now = Date.now()
    let claimed = false
    await runTransaction(db!, async (tx) => {
      const snap = await tx.get(ref)
      const slots: SessionSlot[] = snap.exists() ? (snap.data().slots ?? []) : []
      if (!canClaimSlot(slots, now)) {
        claimed = false
        return
      }
      tx.set(ref, { slots: claimSlot(slots, sessionId, now, deviceLabel) })
      claimed = true
    })
    if (!claimed) return { claim: { ok: false, maxSessions: true }, remoteState: null }

    active = { loginId: id, sessionId, knownDocIds: new Set(), unsubscribes: [], heartbeatTimer: null }
    active.heartbeatTimer = setInterval(() => void heartbeatTick(), 30_000)

    const remoteState = await pullRemoteState(id)
    if (remoteState) active.knownDocIds = new Set((remoteState.createdDocuments ?? []).map((d) => d.id))

    subscribeRemote(id, sessionId, onRemoteUpdate)

    return { claim: { ok: true }, remoteState }
  } catch (e) {
    console.error('startSession: cloud unreachable, continuing offline-only', e)
    return { claim: { ok: true }, remoteState: null }
  }
}

async function heartbeatTick(): Promise<void> {
  if (!active) return
  try {
    const ref = doc(db!, 'sessions', active.loginId)
    await runTransaction(db!, async (tx) => {
      const snap = await tx.get(ref)
      if (!snap.exists()) return
      const slots: SessionSlot[] = snap.data().slots ?? []
      tx.set(ref, { slots: touchSlot(slots, active!.sessionId, Date.now()) })
    })
  } catch (e) {
    console.error('heartbeat failed', e)
  }
}

/** Release this device's session slot. Call on logout and on app quit. */
export async function endSession(): Promise<void> {
  if (!active) return
  const { loginId, sessionId, heartbeatTimer, unsubscribes } = active
  active = null
  if (heartbeatTimer) clearInterval(heartbeatTimer)
  unsubscribes.forEach((u) => u())
  try {
    await ensureAuth()
    const ref = doc(db!, 'sessions', loginId)
    await runTransaction(db!, async (tx) => {
      const snap = await tx.get(ref)
      if (!snap.exists()) return
      const slots: SessionSlot[] = snap.data().slots ?? []
      tx.set(ref, { slots: releaseSlot(slots, sessionId) })
    })
  } catch (e) {
    console.error('endSession: release failed', e)
  }
}

async function pullRemoteState(id: string): Promise<PersistedState | null> {
  try {
    // These 3 reads are independent of one another — fetched in parallel
    // instead of one-after-another so login/sync only pays for the slowest
    // of the 3 round trips, not the sum of all 3.
    const [tablesSnap, miscSnap, docsSnap] = await Promise.all([
      getDoc(doc(db!, 'users', id, 'meta', 'tables')),
      getDoc(doc(db!, 'users', id, 'meta', 'misc')),
      // A Firestore collection has no inherent array order — getDocs()
      // returns documents in Firestore's own arbitrary order (roughly by
      // document ID), not the order the user last arranged them in. Each
      // write tags its document with an explicit `order` index (see
      // pushState) so the user's own ordering (e.g. dragging tiles on Issue
      // Document) survives a push/pull round trip instead of silently
      // resetting.
      getDocs(collection(db!, 'users', id, 'documents'))
    ])
    if (!tablesSnap.exists() && !miscSnap.exists()) return null

    const createdDocuments: CreatedDocument[] = docsSnap.docs
      .map((d) => d.data() as CreatedDocument & { order?: number })
      .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
      .map((v) => ({ id: v.id, name: v.name, docx: v.docx, createdDate: v.createdDate }))

    const t = tablesSnap.exists() ? tablesSnap.data() : {}
    const m = miscSnap.exists() ? miscSnap.data() : {}
    return {
      // The persisted schema version must round-trip through Firebase, or the
      // caller's one-time migrations (e.g. Lakhs -> rupees ECV/Contract Amount)
      // re-run on every launch and compound. Older cloud state has no version
      // field yet — fall back to 1 so those migrations still run exactly once
      // (each is separately guarded against re-inflating already-migrated data).
      version: typeof m.version === 'number' ? m.version : 1,
      tables: t.tables ?? [],
      tablesByOffice: t.tablesByOffice ?? undefined,
      resolution: m.resolution ?? {},
      todos: m.todos ?? [],
      lastGoogleLink: m.lastGoogleLink ?? undefined,
      worksListLinks: m.worksListLinks ?? {},
      tenderReminders: m.tenderReminders ?? [],
      createdDocuments,
      bidDocumentBatches: m.bidDocumentBatches ?? [],
      mbScrutiny: m.mbScrutiny ?? []
    }
  } catch (e) {
    console.error('pullRemoteState failed', e)
    return null
  }
}

function subscribeRemote(
  id: string,
  sessionId: string,
  onRemoteUpdate: (partial: Partial<PersistedState>) => void
): void {
  if (!active) return
  const tablesRef = doc(db!, 'users', id, 'meta', 'tables')
  const miscRef = doc(db!, 'users', id, 'meta', 'misc')
  const docsRef = collection(db!, 'users', id, 'documents')

  // Firestore invokes onSnapshot callbacks from its own internal retry/
  // dispatch timers, outside any request this module's caller made — an
  // exception thrown by the caller's callback here (e.g. touching a
  // BrowserWindow that's since been closed) would otherwise surface as an
  // uncaught exception in the main process and crash the whole app, not
  // just fail this one update.
  const safeUpdate = (partial: Partial<PersistedState>) => {
    try {
      onRemoteUpdate(partial)
    } catch (e) {
      console.error('onRemoteUpdate callback failed', e)
    }
  }

  active.unsubscribes.push(
    onSnapshot(tablesRef, (snap) => {
      // Skip our own not-yet-committed write (local echo) and our own
      // already-committed write (lastWriter === sessionId) — only react to
      // changes that genuinely came from the other device.
      if (snap.metadata.hasPendingWrites) return
      const data = snap.data()
      if (!data || data.lastWriter === sessionId) return
      safeUpdate({ tables: data.tables ?? [], tablesByOffice: data.tablesByOffice ?? undefined })
    })
  )

  active.unsubscribes.push(
    onSnapshot(miscRef, (snap) => {
      if (snap.metadata.hasPendingWrites) return
      const data = snap.data()
      if (!data || data.lastWriter === sessionId) return
      safeUpdate({
        resolution: data.resolution ?? {},
        todos: data.todos ?? [],
        lastGoogleLink: data.lastGoogleLink ?? undefined,
        worksListLinks: data.worksListLinks ?? {},
        tenderReminders: data.tenderReminders ?? [],
        bidDocumentBatches: data.bidDocumentBatches ?? [],
        mbScrutiny: data.mbScrutiny ?? []
      })
    })
  )

  active.unsubscribes.push(
    onSnapshot(docsRef, (snap) => {
      if (snap.metadata.hasPendingWrites) return
      // A per-doc lastWriter filter here would risk dropping a genuine
      // remote change whose sibling doc happens to echo our own recent
      // write in the same snapshot. The collection is small, so just
      // re-pull the full set instead of reconstructing it from the diff.
      void pullRemoteState(id).then((full) => {
        if (full) safeUpdate({ createdDocuments: full.createdDocuments })
      })
    })
  )
}

/** Push the current workspace to the cloud, tagged as written by this device. */
export async function pushState(state: PersistedState): Promise<void> {
  if (!active) return
  const { loginId, sessionId } = active
  try {
    await ensureAuth()
    // Update ONLY the office this session is on (state.currentOfficeKey) inside
    // the tablesByOffice map, merging so a concurrent session working on another
    // office keeps its own entry. `tables` stays for backward-compat with older
    // clients (the current office's data). `merge: true` deep-merges the map.
    const officeK = state.currentOfficeKey
    const officeEntry = officeK ? state.tablesByOffice?.[officeK] ?? state.tables ?? [] : null
    await setDoc(
      doc(db!, 'users', loginId, 'meta', 'tables'),
      {
        tables: state.tables ?? [],
        ...(officeK ? { tablesByOffice: { [officeK]: officeEntry } } : {}),
        updatedAt: serverTimestamp(),
        lastWriter: sessionId
      },
      { merge: true }
    )
    await setDoc(doc(db!, 'users', loginId, 'meta', 'misc'), {
      // Round-trip the schema version so one-time migrations run only once (see
      // pullRemoteState) instead of re-running — and compounding — every launch.
      version: state.version,
      resolution: state.resolution ?? {},
      todos: state.todos ?? [],
      lastGoogleLink: state.lastGoogleLink ?? null,
      worksListLinks: state.worksListLinks ?? {},
      tenderReminders: state.tenderReminders ?? [],
      bidDocumentBatches: state.bidDocumentBatches ?? [],
      mbScrutiny: state.mbScrutiny ?? [],
      updatedAt: serverTimestamp(),
      lastWriter: sessionId
    })

    const currentIds = new Set((state.createdDocuments ?? []).map((d) => d.id))
    const orderedDocs = state.createdDocuments ?? []
    for (let i = 0; i < orderedDocs.length; i++) {
      const d = orderedDocs[i]
      try {
        // Firestore caps a single document at 1 MiB — a base64 .docx is
        // larger than the plain HTML this field used to hold, though real
        // OOXML is typically more compact than Word's own verbose clipboard
        // HTML to begin with. An unusually large pasted document (many
        // pages, embedded images) could still approach that ceiling —
        // caught per-document so one oversized document doesn't also block
        // every other document (and the cleanup step below) from syncing.
        await setDoc(doc(db!, 'users', loginId, 'documents', d.id), {
          id: d.id,
          name: d.name,
          docx: d.docx,
          createdDate: d.createdDate,
          order: i,
          updatedAt: serverTimestamp(),
          lastWriter: sessionId
        })
      } catch (e) {
        console.error(`pushState: failed to sync document "${d.name}"`, e)
      }
    }
    if (active) {
      for (const oldId of active.knownDocIds) {
        if (!currentIds.has(oldId)) await deleteDoc(doc(db!, 'users', loginId, 'documents', oldId))
      }
      active.knownDocIds = currentIds
    }
  } catch (e) {
    console.error('pushState failed', e)
  }
}

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
import { canClaimSlot, claimSlot, releaseSlot, touchSlot, clearSlots, liveSlots, type SessionSlot } from '../core/sessionSlots'
import type { PersistedState, CreatedDocument, TodoItem, MBScrutinyItem } from '../core/types'
import { LEGACY_SEED_DOC_IDS, pruneLegacyDocuments } from '../core/types'

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

export function normalizeId(loginId: string): string {
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
  /** Todo ids this session last saw in the cloud — so a delete removes only a
   * todo it knew about, never one another device added since (see pushState). */
  knownTodoIds: Set<string>
  /** MB Scrutiny record ids this session last saw in the cloud — same per-item
   * delete safety as knownTodoIds (see pushState). */
  knownMbIds: Set<string>
  /**
   * Signature (order|name|docx) of each document as last written to the cloud,
   * so pushState re-uploads only the documents that actually changed instead of
   * every ~100KB base64 .docx on every autosave — the biggest cloud-push cost.
   */
  pushedDocSig: Map<string, string>
  unsubscribes: Unsubscribe[]
  heartbeatTimer: ReturnType<typeof setInterval> | null
}

let active: ActiveSession | null = null

/**
 * Call after a successful password check. Claims one of the session slots
 * for this login ID, pulls whatever's already synced for it, and starts a
 * live listener for changes made from the other device.
 *
 * If the cloud is unreachable, the login proceeds anyway (offline-only, no
 * sync, no session cap) rather than locking the user out over a network hiccup.
 *
 * `force`: skip the slot check and sign every other device out first — the
 * login screen's "log out other devices" recovery once all
 * MAX_CONCURRENT_SESSIONS are taken. Only reachable after the password has
 * already been re-verified (see electron/main.ts's login handler), since
 * this immediately ends every other device's session.
 */
export async function startSession(
  loginId: string,
  deviceLabel: string,
  onRemoteUpdate: (partial: Partial<PersistedState>) => void,
  force = false
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
      const existing: SessionSlot[] = snap.exists() ? (snap.data().slots ?? []) : []
      const slots = force ? clearSlots() : existing
      if (!canClaimSlot(slots, now)) {
        claimed = false
        return
      }
      tx.set(ref, { slots: claimSlot(slots, sessionId, now, deviceLabel) })
      claimed = true
    })
    if (!claimed) return { claim: { ok: false, maxSessions: true }, remoteState: null }

    active = { loginId: id, sessionId, knownDocIds: new Set(), knownTodoIds: new Set(), knownMbIds: new Set(), pushedDocSig: new Map(), unsubscribes: [], heartbeatTimer: null }
    active.heartbeatTimer = setInterval(() => void heartbeatTick(), 120_000)

    const remoteState = await pullRemoteState(id)
    if (remoteState) {
      active.knownDocIds = new Set((remoteState.createdDocuments ?? []).map((d) => d.id))
      active.knownTodoIds = new Set((remoteState.todos ?? []).map((t) => t.id))
      active.knownMbIds = new Set((remoteState.mbScrutiny ?? []).map((it) => it.id))
      // Seed the doc signatures from what the cloud already has, so the first
      // push after login doesn't needlessly re-upload documents that match.
      active.pushedDocSig = new Map(
        (remoteState.createdDocuments ?? []).map((d, i) => [d.id, `${i}|${d.name}|${d.docx}`])
      )
    }

    // One-time cloud cleanup: permanently delete the superseded legacy example
    // documents from Firestore (best-effort, fire-and-forget). pullRemoteState
    // already hides them from the UI; this removes the underlying records so
    // they don't linger for other devices. Idempotent — deleting an absent doc
    // is a harmless no-op — so after the first login it costs a few tiny writes.
    if (remoteState) {
      void Promise.allSettled(
        LEGACY_SEED_DOC_IDS.map((legacyId) => deleteDoc(doc(db!, 'users', id, 'documents', legacyId)))
      )
    }

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

/** This device's own session id, or null if not signed in. Used to mark "this device" in Settings' Active Devices list. */
export function currentSessionId(): string | null {
  return active?.sessionId ?? null
}

/** Lists every live session for a login — Settings' "Active Devices" panel. Empty (not thrown) if the cloud is unreachable, matching this module's offline-first fallback elsewhere. */
export async function listSessions(loginId: string): Promise<SessionSlot[]> {
  try {
    await ensureAuth()
    const ref = doc(db!, 'sessions', normalizeId(loginId))
    const snap = await getDoc(ref)
    const slots: SessionSlot[] = snap.exists() ? (snap.data().slots ?? []) : []
    return liveSlots(slots, Date.now())
  } catch (e) {
    console.error('listSessions failed', e)
    return []
  }
}

/**
 * Ends one specific OTHER device's session — Settings' "Active Devices" panel
 * kicking a stale/unwanted session. Never touches this device's own local
 * `active` state (that's endSession's job, via the normal logout flow) — a
 * sessionId that happens to be this device's own is simply released same as
 * any other, but electron/main.ts's handler never passes one, so in practice
 * this only ever removes a slot no longer backed by a running app.
 */
export async function endOtherSession(loginId: string, sessionId: string): Promise<void> {
  await ensureAuth()
  const ref = doc(db!, 'sessions', normalizeId(loginId))
  await runTransaction(db!, async (tx) => {
    const snap = await tx.get(ref)
    if (!snap.exists()) return
    const slots: SessionSlot[] = snap.data().slots ?? []
    tx.set(ref, { slots: releaseSlot(slots, sessionId) })
  })
}

async function pullRemoteState(id: string): Promise<PersistedState | null> {
  try {
    // These reads are independent of one another — fetched in parallel
    // instead of one-after-another so login/sync only pays for the slowest
    // round trip, not the sum of all of them.
    const [tablesSnap, miscSnap, docsSnap, todosSnap, mbSnap] = await Promise.all([
      getDoc(doc(db!, 'users', id, 'meta', 'tables')),
      getDoc(doc(db!, 'users', id, 'meta', 'misc')),
      // A Firestore collection has no inherent array order — getDocs()
      // returns documents in Firestore's own arbitrary order (roughly by
      // document ID), not the order the user last arranged them in. Each
      // write tags its document with an explicit `order` index (see
      // pushState) so the user's own ordering (e.g. dragging tiles on Issue
      // Document) survives a push/pull round trip instead of silently
      // resetting.
      getDocs(collection(db!, 'users', id, 'documents')),
      // Todos sync per-item (like documents), so one device's stale whole-blob
      // write can't wipe a task another device just added. `order` preserves the
      // user's arrangement across a round trip.
      getDocs(collection(db!, 'users', id, 'todos')),
      // MB Scrutiny records sync per-item for the same reason — the whole-blob
      // misc.mbScrutiny write used to let a concurrent session's stale copy wipe
      // a record another device had just added. `order` keeps receipt order.
      getDocs(collection(db!, 'users', id, 'mbScrutiny'))
    ])
    if (!tablesSnap.exists() && !miscSnap.exists()) return null

    // Strip the superseded legacy example documents here so a cloud collection
    // that still holds them (an older workspace whose local prune hasn't yet
    // pushed) can never re-surface them on screen via the live snapshot. The
    // physical Firestore records are cleaned up separately — see startSession.
    const createdDocuments: CreatedDocument[] = pruneLegacyDocuments(
      docsSnap.docs
        .map((d) => d.data() as CreatedDocument & { order?: number })
        .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
        .map((v) => ({ id: v.id, name: v.name, docx: v.docx, createdDate: v.createdDate }))
    )

    const t = tablesSnap.exists() ? tablesSnap.data() : {}
    const m = miscSnap.exists() ? miscSnap.data() : {}
    // Prefer the per-item todos collection; fall back to the legacy misc.todos
    // blob for a workspace that predates per-item sync (migrated up on next push).
    const todosFromCollection: TodoItem[] = todosSnap.docs
      .map((d) => d.data() as TodoItem & { order?: number })
      .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
      .map((v) => ({
        id: v.id,
        text: v.text,
        createdDate: v.createdDate,
        targetDate: v.targetDate,
        done: v.done,
        ...(v.completedDate ? { completedDate: v.completedDate } : {}),
        ...(v.officeKey ? { officeKey: v.officeKey } : {})
      }))
    const todos = todosFromCollection.length > 0 ? todosFromCollection : (m.todos ?? [])
    // Same as todos: prefer the per-item mbScrutiny collection, falling back to
    // the legacy misc.mbScrutiny blob for a workspace that predates per-item
    // sync (migrated up on next push). Firestore rejects `undefined`, so the
    // optional fields were only written when present — re-spread just those.
    const mbFromCollection: MBScrutinyItem[] = mbSnap.docs
      .map((d) => d.data() as MBScrutinyItem & { order?: number })
      .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
      .map((v) => ({
        id: v.id,
        serialNo: v.serialNo,
        mbNo: v.mbNo,
        agencyName: v.agencyName,
        receivedDate: v.receivedDate,
        done: v.done,
        ...(v.targetDate ? { targetDate: v.targetDate } : {}),
        ...(v.completedDate ? { completedDate: v.completedDate } : {}),
        ...(v.remarks ? { remarks: v.remarks } : {}),
        ...(v.officeKey ? { officeKey: v.officeKey } : {})
      }))
    const mbScrutiny = mbFromCollection.length > 0 ? mbFromCollection : (m.mbScrutiny ?? [])
    return {
      // The persisted schema version must round-trip through Firebase, or the
      // caller's one-time migrations (e.g. Lakhs -> rupees ECV/Contract Amount)
      // re-run on every launch and compound. Older cloud state has no version
      // field yet — fall back to 1 so those migrations still run exactly once
      // (each is separately guarded against re-inflating already-migrated data).
      version: typeof m.version === 'number' ? m.version : 1,
      tables: t.tables ?? [],
      tablesByOffice: t.tablesByOffice ?? undefined,
      // Restores the Head Draughtsman's chosen office on the next login / a
      // fresh machine — see PersistedState.office's own doc comment. Was
      // missing here entirely (and in pushState's write above) until now, so
      // every login silently lost it even though it was saved locally.
      office: m.office ?? undefined,
      resolution: m.resolution ?? {},
      todos,
      lastGoogleLink: m.lastGoogleLink ?? undefined,
      worksListLinks: m.worksListLinks ?? {},
      qcParties: m.qcParties ?? undefined,
      tenderReminders: m.tenderReminders ?? [],
      createdDocuments,
      bidDocumentBatches: m.bidDocumentBatches ?? [],
      mbScrutiny
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
  const todosRef = collection(db!, 'users', id, 'todos')
  const mbRef = collection(db!, 'users', id, 'mbScrutiny')

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
      // NOTE: todos and mbScrutiny are intentionally NOT read from the misc blob
      // here — they sync via their own per-item collections below, so a stale
      // whole-blob misc write can't clobber either list.
      safeUpdate({
        resolution: data.resolution ?? {},
        lastGoogleLink: data.lastGoogleLink ?? undefined,
        worksListLinks: data.worksListLinks ?? {},
        qcParties: data.qcParties ?? undefined,
        tenderReminders: data.tenderReminders ?? [],
        bidDocumentBatches: data.bidDocumentBatches ?? []
      })
    })
  )

  // The direct pullRemoteState in startSession already loaded the current docs
  // and todos, so the FIRST snapshot each collection listener delivers is just
  // that same state echoed back. Reacting to it would fire a redundant full
  // pullRemoteState (2 extra cloud pulls per login, ~half of a login's total
  // reads) for no new data. Skip that initial delivery and only react to
  // genuine later changes from the other device.
  let firstTodosSnap = true
  let firstDocsSnap = true
  let firstMbSnap = true

  active.unsubscribes.push(
    onSnapshot(todosRef, (snap) => {
      if (snap.metadata.hasPendingWrites) return
      if (firstTodosSnap) {
        firstTodosSnap = false
        return
      }
      // Re-pull the whole (small) set rather than reconstruct from the diff —
      // same reasoning as documents. Never let an EMPTY cloud collection wipe
      // the on-screen list: a workspace still on the legacy misc.todos blob has
      // an empty todos collection until its first push migrates it, and applying
      // [] here would erase those tasks. The migrated tasks get pushed up moments
      // later; a genuine "deleted every task" reflects on the next full pull.
      void pullRemoteState(id).then((full) => {
        if (full && (full.todos?.length ?? 0) > 0) safeUpdate({ todos: full.todos })
      })
    })
  )

  active.unsubscribes.push(
    onSnapshot(docsRef, (snap) => {
      if (snap.metadata.hasPendingWrites) return
      if (firstDocsSnap) {
        firstDocsSnap = false
        return
      }
      // A per-doc lastWriter filter here would risk dropping a genuine
      // remote change whose sibling doc happens to echo our own recent
      // write in the same snapshot. The collection is small, so just
      // re-pull the full set instead of reconstructing it from the diff.
      void pullRemoteState(id).then((full) => {
        // Never let an EMPTY cloud documents collection overwrite what's on
        // screen. On a fresh system the app injects the bundled default
        // documents (and seeds the examples) during load, then this snapshot
        // fires; if the cloud collection is still empty (a workspace whose
        // works data synced but whose documents never did) sending [] here
        // would wipe those freshly-injected defaults — and because the local
        // seededDocVersion is now current, the injection never re-runs, so the
        // Tools/Issue-Documents tabs stay permanently blank. The just-injected
        // documents get pushed up moments later, after which real snapshots
        // carry them. A genuine "user deleted every document" is not worth
        // reintroducing that whole-workspace wipe for.
        if (full && (full.createdDocuments?.length ?? 0) > 0) safeUpdate({ createdDocuments: full.createdDocuments })
      })
    })
  )

  active.unsubscribes.push(
    onSnapshot(mbRef, (snap) => {
      if (snap.metadata.hasPendingWrites) return
      if (firstMbSnap) {
        firstMbSnap = false
        return
      }
      // Re-pull the whole (small) set rather than reconstruct from the diff —
      // same reasoning as todos. Never let an EMPTY cloud collection wipe the
      // on-screen list: a workspace still on the legacy misc.mbScrutiny blob has
      // an empty collection until its first push migrates it, and applying []
      // here would erase those records. The migrated records get pushed up
      // moments later; a genuine "deleted every MB" reflects on the next full pull.
      void pullRemoteState(id).then((full) => {
        if (full && (full.mbScrutiny?.length ?? 0) > 0) safeUpdate({ mbScrutiny: full.mbScrutiny })
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
    await setDoc(
      doc(db!, 'users', loginId, 'meta', 'misc'),
      {
        // Round-trip the schema version so one-time migrations run only once (see
        // pullRemoteState) instead of re-running — and compounding — every launch.
        version: state.version,
        resolution: state.resolution ?? {},
        todos: state.todos ?? [],
        lastGoogleLink: state.lastGoogleLink ?? null,
        worksListLinks: state.worksListLinks ?? {},
        qcParties: state.qcParties ?? {},
        tenderReminders: state.tenderReminders ?? [],
        bidDocumentBatches: state.bidDocumentBatches ?? [],
        mbScrutiny: state.mbScrutiny ?? [],
        // Deliberately omitted (not written as null) when not yet chosen —
        // `merge: true` below then leaves whatever office is already stored in
        // the cloud untouched, matching PersistedState.office's own LWW
        // contract ("never let an empty office overwrite a good one already
        // stored"). Every other field above is unconditional/defaulted, so
        // switching this write to merge:true doesn't change their behavior.
        ...(state.office ? { office: state.office } : {}),
        updatedAt: serverTimestamp(),
        lastWriter: sessionId
      },
      { merge: true }
    )

    const currentIds = new Set((state.createdDocuments ?? []).map((d) => d.id))
    const orderedDocs = state.createdDocuments ?? []
    for (let i = 0; i < orderedDocs.length; i++) {
      const d = orderedDocs[i]
      // Skip documents that haven't changed since we last pushed them (same
      // order, name and .docx content). The documents are the same 9-11 default
      // templates on nearly every save, so this turns a ~1MB upload on every
      // autosave into (usually) zero document writes — only the edited data
      // (tables/todos) actually syncs. A reorder or template refresh changes the
      // signature and re-uploads just that doc.
      const sig = `${i}|${d.name}|${d.docx}`
      if (active?.pushedDocSig.get(d.id) === sig) continue
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
        active?.pushedDocSig.set(d.id, sig)
      } catch (e) {
        console.error(`pushState: failed to sync document "${d.name}"`, e)
      }
    }
    if (active) {
      for (const oldId of active.knownDocIds) {
        if (!currentIds.has(oldId)) {
          await deleteDoc(doc(db!, 'users', loginId, 'documents', oldId))
          active.pushedDocSig.delete(oldId)
        }
      }
      active.knownDocIds = currentIds
    }

    // Todos: one Firestore doc per task (mirrors documents), so a stale
    // whole-blob push can't wipe a task another device just added. Firestore
    // rejects `undefined`, so completedDate is only written when present. The
    // legacy misc.todos above stays dual-written for not-yet-updated clients.
    const orderedTodos = state.todos ?? []
    const currentTodoIds = new Set(orderedTodos.map((t) => t.id))
    for (let i = 0; i < orderedTodos.length; i++) {
      const t = orderedTodos[i]
      try {
        await setDoc(doc(db!, 'users', loginId, 'todos', t.id), {
          id: t.id,
          text: t.text,
          createdDate: t.createdDate,
          targetDate: t.targetDate,
          done: t.done,
          ...(t.completedDate ? { completedDate: t.completedDate } : {}),
          ...(t.officeKey ? { officeKey: t.officeKey } : {}),
          order: i,
          updatedAt: serverTimestamp(),
          lastWriter: sessionId
        })
      } catch (e) {
        console.error(`pushState: failed to sync todo "${t.text}"`, e)
      }
    }
    if (active) {
      for (const oldId of active.knownTodoIds) {
        if (!currentTodoIds.has(oldId)) await deleteDoc(doc(db!, 'users', loginId, 'todos', oldId))
      }
      active.knownTodoIds = currentTodoIds
    }

    // MB Scrutiny: one Firestore doc per record (mirrors todos), so a stale
    // whole-blob push can't wipe a record another device just added. Firestore
    // rejects `undefined`, so the optional fields are only written when present.
    // The legacy misc.mbScrutiny above stays dual-written for not-yet-updated
    // clients.
    const orderedMb = state.mbScrutiny ?? []
    const currentMbIds = new Set(orderedMb.map((it) => it.id))
    for (let i = 0; i < orderedMb.length; i++) {
      const it = orderedMb[i]
      try {
        await setDoc(doc(db!, 'users', loginId, 'mbScrutiny', it.id), {
          id: it.id,
          serialNo: it.serialNo,
          mbNo: it.mbNo,
          agencyName: it.agencyName,
          receivedDate: it.receivedDate,
          done: it.done,
          ...(it.targetDate ? { targetDate: it.targetDate } : {}),
          ...(it.completedDate ? { completedDate: it.completedDate } : {}),
          ...(it.remarks ? { remarks: it.remarks } : {}),
          ...(it.officeKey ? { officeKey: it.officeKey } : {}),
          order: i,
          updatedAt: serverTimestamp(),
          lastWriter: sessionId
        })
      } catch (e) {
        console.error(`pushState: failed to sync MB record "${it.mbNo}"`, e)
      }
    }
    if (active) {
      for (const oldId of active.knownMbIds) {
        if (!currentMbIds.has(oldId)) await deleteDoc(doc(db!, 'users', loginId, 'mbScrutiny', oldId))
      }
      active.knownMbIds = currentMbIds
    }
  } catch (e) {
    console.error('pushState failed', e)
  }
}

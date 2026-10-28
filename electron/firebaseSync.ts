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
  deleteField,
  deleteDoc,
  runTransaction,
  collection,
  onSnapshot,
  serverTimestamp,
  type Firestore,
  type QuerySnapshot,
  type Unsubscribe
} from 'firebase/firestore'
import { canClaimSlot, claimSlot, releaseSlot, touchSlot, clearSlots, liveSlots, HEARTBEAT_MS, type SessionSlot } from '../core/sessionSlots'
import type { PersistedState, CreatedDocument, TodoItem, MBScrutinyItem, PendingNoticeItem, TenderConditionsEntry, ShareResult } from '../core/types'
import { cleanTenderConditionsForCloud } from '../core/tenderConditions'
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
  /** Non-conclusion Notice ids this session last saw in the cloud — same per-item delete safety (see pushState). */
  knownNoticeIds: Set<string>
  /**
   * Signature (order|name|docx) of each document as last written to the cloud,
   * so pushState re-uploads only the documents that actually changed instead of
   * every ~100KB base64 .docx on every autosave — the biggest cloud-push cost.
   */
  pushedDocSig: Map<string, string>
  /** JSON of each office's Works List as last seen in / written to the cloud, so a push writes only the offices that changed here (never re-sends — and so overwrites — an office another device edited). */
  pushedWorksSig: Map<string, string>
  /** Signature of each To Do / MB record as last seen in / written to the cloud, so a push sends only the items that changed here — a stale device can no longer overwrite items another device edited. */
  pushedTodoSig: Map<string, string>
  pushedMbSig: Map<string, string>
  pushedNoticeSig: Map<string, string>
  /** Signature of the shared tables / misc documents as last written from here, so an autosave that changed nothing in them skips both writes (each is a billed write, and a stale rewrite could clobber another device's edit). */
  pushedTablesSig: string
  pushedMiscSig: string
  unsubscribes: Unsubscribe[]
  heartbeatTimer: ReturnType<typeof setInterval> | null
}

let active: ActiveSession | null = null
/** Saved-tender IDs the cloud is known to hold. The misc write merges maps, so a removed ID must be deleted explicitly. */
let knownCloudTenderIds = new Set<string>()
/** Office keys the cloud holds a Monitoring Format summary for. The misc write merges maps, so a deleted office's summary must be removed with deleteField(). */
let knownCloudMfKeys = new Set<string>()
/** Why the last pushState failed (shown by the Sync button instead of a vague "check connection"). */
let lastPushError = ''
export function getLastPushError(): string {
  return lastPushError
}

/**
 * Works Lists live one Firestore document PER OFFICE (users/<id>/worksByOffice/<officeKey>). They used to
 * share a single meta/tables document, which grew past Firestore's 1 MiB per-document cap once several
 * offices were synced — after which every push was rejected.
 */
const worksDocId = (officeKey: string): string => encodeURIComponent(officeKey).replace(/\./g, '%2E')

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

    active = { loginId: id, sessionId, knownDocIds: new Set(), knownTodoIds: new Set(), knownMbIds: new Set(), knownNoticeIds: new Set(), pushedDocSig: new Map(), pushedWorksSig: new Map(), pushedTodoSig: new Map(), pushedMbSig: new Map(), pushedNoticeSig: new Map(), pushedTablesSig: '', pushedMiscSig: '', unsubscribes: [], heartbeatTimer: null }
    active.heartbeatTimer = setInterval(() => void heartbeatTick(), HEARTBEAT_MS)

    const remoteState = await pullRemoteState(id)
    if (remoteState) {
      active.knownDocIds = new Set((remoteState.createdDocuments ?? []).map((d) => d.id))
      active.knownTodoIds = new Set((remoteState.todos ?? []).map((t) => t.id))
      active.knownMbIds = new Set((remoteState.mbScrutiny ?? []).map((it) => it.id))
      active.knownNoticeIds = new Set((remoteState.pendingNotices ?? []).map((it) => it.id))
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

/** The cloud fields of a To Do (without the write timestamp), shared by the push and by the pull's signature so both agree. */
function todoFields(t: TodoItem, order: number) {
  return {
    id: t.id,
    text: t.text,
    createdDate: t.createdDate,
    targetDate: t.targetDate,
    done: t.done,
    ...(t.completedDate ? { completedDate: t.completedDate } : {}),
    ...(t.officeKey ? { officeKey: t.officeKey } : {}),
    order
  }
}

/** Same for an MB Scrutiny record. Firestore rejects `undefined`, so optional fields are only present when set. */
function mbFields(it: MBScrutinyItem, order: number) {
  return {
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
    order
  }
}

/** A Non-conclusion Notice as stored in its own Firestore doc. Round-tripped through JSON so Firestore never sees `undefined`. */
function noticeFields(it: PendingNoticeItem, order: number) {
  return { ...(JSON.parse(JSON.stringify(it)) as PendingNoticeItem), order }
}

type DocLike = { data(): unknown }

function noticesFromDocs(docs: DocLike[]): PendingNoticeItem[] {
  return docs
    .map((d) => d.data() as PendingNoticeItem & { order?: number; updatedAt?: unknown; lastWriter?: unknown })
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
    .map(({ order: _o, updatedAt: _u, lastWriter: _l, ...item }) => item as PendingNoticeItem)
}

function todosFromDocs(docs: DocLike[]): TodoItem[] {
  return docs
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
}

function mbFromDocs(docs: DocLike[]): MBScrutinyItem[] {
  return docs
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
}

function createdDocsFromDocs(docs: DocLike[]): CreatedDocument[] {
  return pruneLegacyDocuments(
    docs
      .map((d) => d.data() as CreatedDocument & { order?: number })
      .sort((a, b) => (a.order ?? 0) - (b.order ?? 0))
      .map((v) => ({ id: v.id, name: v.name, docx: v.docx, createdDate: v.createdDate }))
  )
}

async function pullRemoteState(id: string): Promise<PersistedState | null> {
  try {
    // These reads are independent of one another — fetched in parallel
    // instead of one-after-another so login/sync only pays for the slowest
    // round trip, not the sum of all of them.
    const [tablesSnap, miscSnap, docsSnap, todosSnap, mbSnap, noticesSnap, worksSnap] = await Promise.all([
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
      getDocs(collection(db!, 'users', id, 'mbScrutiny')),
      // Notices sync per-item too: a bulk run of notices on one device must not be wiped by another device's stale whole-blob write.
      getDocs(collection(db!, 'users', id, 'pendingNotices')),
      // One document per office's Works List (see worksDocId).
      getDocs(collection(db!, 'users', id, 'worksByOffice'))
    ])
    if (!tablesSnap.exists() && !miscSnap.exists() && worksSnap.empty) return null

    // Strip the superseded legacy example documents here so a cloud collection
    // that still holds them (an older workspace whose local prune hasn't yet
    // pushed) can never re-surface them on screen via the live snapshot. The
    // physical Firestore records are cleaned up separately — see startSession.
    const createdDocuments: CreatedDocument[] = createdDocsFromDocs(docsSnap.docs)

    const t = tablesSnap.exists() ? tablesSnap.data() : {}
    // Offices from the legacy shared document first, then the per-office documents on top (they're newer).
    const worksByOffice: Record<string, unknown[]> = { ...((t.tablesByOffice as Record<string, unknown[]> | undefined) ?? {}) }
    if (active) active.pushedWorksSig = new Map()
    for (const d of worksSnap.docs) {
      const data = d.data() as { officeKey?: string; tables?: unknown[] }
      if (typeof data.officeKey !== 'string') continue
      worksByOffice[data.officeKey] = data.tables ?? []
      active?.pushedWorksSig.set(data.officeKey, JSON.stringify(data.tables ?? []))
    }
    const m = miscSnap.exists() ? miscSnap.data() : {}
    knownCloudTenderIds = new Set(Object.keys((m.tenderConditions as Record<string, unknown> | undefined) ?? {}))
    knownCloudMfKeys = new Set(Object.keys((m.monitoringFormatByOffice as Record<string, unknown> | undefined) ?? {}))
    // Prefer the per-item todos collection; fall back to the legacy misc.todos
    // blob for a workspace that predates per-item sync (migrated up on next push).
    const todosFromCollection: TodoItem[] = todosFromDocs(todosSnap.docs)
    const todos = todosFromCollection.length > 0 ? todosFromCollection : (m.todos ?? [])
    // Same as todos: prefer the per-item mbScrutiny collection, falling back to
    // the legacy misc.mbScrutiny blob for a workspace that predates per-item
    // sync (migrated up on next push). Firestore rejects `undefined`, so the
    // optional fields were only written when present — re-spread just those.
    const mbFromCollection: MBScrutinyItem[] = mbFromDocs(mbSnap.docs)
    // Notices: prefer the per-item collection, falling back to the legacy misc.pendingNotices blob for a workspace that predates it (migrated up on next push).
    const noticesFromCollection: PendingNoticeItem[] = noticesFromDocs(noticesSnap.docs)
    const pendingNotices = noticesFromCollection.length > 0 ? noticesFromCollection : (m.pendingNotices ?? [])
    const mbScrutiny = mbFromCollection.length > 0 ? mbFromCollection : (m.mbScrutiny ?? [])
    // What the cloud holds right now — the next push skips any item still identical to this.
    if (active) {
      // Cloud content just read: the next push must write the shared documents again rather than trust a stale signature.
      active.pushedTablesSig = ''
      active.pushedMiscSig = ''
      active.pushedTodoSig = new Map(todosFromCollection.map((v, i) => [v.id, JSON.stringify(todoFields(v, i))]))
      active.pushedMbSig = new Map(mbFromCollection.map((v, i) => [v.id, JSON.stringify(mbFields(v, i))]))
      active.pushedNoticeSig = new Map(noticesFromCollection.map((v, i) => [v.id, JSON.stringify(noticeFields(v, i))]))
    }
    return {
      // The persisted schema version must round-trip through Firebase, or the
      // caller's one-time migrations (e.g. Lakhs -> rupees ECV/Contract Amount)
      // re-run on every launch and compound. Older cloud state has no version
      // field yet — fall back to 1 so those migrations still run exactly once
      // (each is separately guarded against re-inflating already-migrated data).
      version: typeof m.version === 'number' ? m.version : 1,
      tables: t.tables ?? [],
      tablesByOffice: Object.keys(worksByOffice).length > 0 ? (worksByOffice as PersistedState['tablesByOffice']) : undefined,
      // Restores the Head Draughtsman's chosen office on the next login / a
      // fresh machine — see PersistedState.office's own doc comment. Was
      // missing here entirely (and in pushState's write above) until now, so
      // every login silently lost it even though it was saved locally.
      office: m.office ?? undefined,
      resolution: m.resolution ?? {},
      todos,
      lastGoogleLink: m.lastGoogleLink ?? undefined,
      worksListLinks: m.worksListLinks ?? {},
      contacts: m.contacts ?? undefined,
      qcParties: m.qcParties ?? undefined,
      monitoringFormatByOffice: m.monitoringFormatByOffice ?? undefined,
      monitoringFormatLinks: m.monitoringFormatLinks ?? {},
      tenderConditions: m.tenderConditions ?? undefined,
      tenderReminders: m.tenderReminders ?? [],
      pendingNotices,
      createdDocuments,
      bidDocumentBatches: m.bidDocumentBatches ?? [],
      mbScrutiny
    }
  } catch (e) {
    console.error('pullRemoteState failed', e)
    return null
  }
}

/**
 * Whether a collection snapshot carries a change another device made. Decided per document, not by
 * snap.metadata.hasPendingWrites: that flag is set whenever this device has ANY write in flight, so a
 * snapshot mixing this device's own write with another device's change used to be thrown away whole — and
 * Firestore never re-delivers it, so the other device's change simply never appeared. A removal carries only
 * the deleted document's last data, so it counts unless it is this device's own still-pending delete.
 */
function hasRemoteChange(snap: QuerySnapshot, sessionId: string): boolean {
  return snap.docChanges().some((ch) =>
    ch.type === 'removed'
      ? !snap.metadata.hasPendingWrites
      : !ch.doc.metadata.hasPendingWrites && ch.doc.data().lastWriter !== sessionId
  )
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
  const noticesRef = collection(db!, 'users', id, 'pendingNotices')

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
      // Once Works Lists moved to their own per-office documents this one only carries the flat legacy copy;
      // applying it would put another device's current office on screen. Older clients still use it.
      if (data.perOffice) return
      if (active) active.pushedTablesSig = ''
      safeUpdate({ tables: data.tables ?? [], tablesByOffice: data.tablesByOffice ?? undefined })
    })
  )

  let firstWorksSnap = true
  active.unsubscribes.push(
    onSnapshot(collection(db!, 'users', id, 'worksByOffice'), (snap) => {
      // Not dropped on snap.metadata.hasPendingWrites: another device's change can arrive in the same snapshot as
      // this device's own in-flight write, and no later snapshot repeats it. Own writes are filtered per document.
      if (firstWorksSnap) {
        firstWorksSnap = false
        return
      }
      const incoming: Record<string, unknown[]> = {}
      for (const ch of snap.docChanges()) {
        if (ch.type === 'removed' || ch.doc.metadata.hasPendingWrites) continue
        const data = ch.doc.data() as { officeKey?: string; tables?: unknown[]; lastWriter?: string }
        if (typeof data.officeKey !== 'string' || data.lastWriter === sessionId) continue
        incoming[data.officeKey] = data.tables ?? []
        active?.pushedWorksSig.set(data.officeKey, JSON.stringify(data.tables ?? []))
      }
      if (Object.keys(incoming).length > 0) safeUpdate({ tablesByOffice: incoming as PersistedState['tablesByOffice'] })
    })
  )

  active.unsubscribes.push(
    onSnapshot(miscRef, (snap) => {
      if (snap.metadata.hasPendingWrites) return
      const data = snap.data()
      if (!data || data.lastWriter === sessionId) return
      // NOTE: todos, mbScrutiny and pendingNotices are intentionally NOT read from the misc blob
      // here — they sync via their own per-item collections below, so a stale
      // whole-blob misc write can't clobber either list.
      knownCloudTenderIds = new Set(Object.keys((data.tenderConditions as Record<string, unknown> | undefined) ?? {}))
      knownCloudMfKeys = new Set(Object.keys((data.monitoringFormatByOffice as Record<string, unknown> | undefined) ?? {}))
      // Another device changed this document: a later local change back to what we last wrote must still be pushed.
      if (active) active.pushedMiscSig = ''
      safeUpdate({
        resolution: data.resolution ?? {},
        lastGoogleLink: data.lastGoogleLink ?? undefined,
        worksListLinks: data.worksListLinks ?? {},
        contacts: data.contacts ?? undefined,
        qcParties: data.qcParties ?? undefined,
        monitoringFormatByOffice: data.monitoringFormatByOffice ?? undefined,
        monitoringFormatLinks: data.monitoringFormatLinks ?? {},
        // An explicit {} (not undefined) so a tender deleted on the other device is removed here too.
        tenderConditions: data.tenderConditions ?? {},
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
      if (firstTodosSnap) {
        firstTodosSnap = false
        return
      }
      if (!hasRemoteChange(snap, sessionId)) return
      // The snapshot already holds the whole current set (the server only sends the changed documents), so the
      // list is built from it instead of re-reading every document — a full re-read cost a read per task on every
      // change made on another device. Never let an EMPTY cloud collection wipe
      // the on-screen list: a workspace still on the legacy misc.todos blob has
      // an empty todos collection until its first push migrates it, and applying
      // [] here would erase those tasks. The migrated tasks get pushed up moments
      // later; a genuine "deleted every task" reflects on the next full pull.
      const todos = todosFromDocs(snap.docs)
      if (todos.length > 0) {
        if (active) active.pushedTodoSig = new Map(todos.map((v, i) => [v.id, JSON.stringify(todoFields(v, i))]))
        safeUpdate({ todos })
      }
    })
  )

  active.unsubscribes.push(
    onSnapshot(docsRef, (snap) => {
      if (firstDocsSnap) {
        firstDocsSnap = false
        return
      }
      if (!hasRemoteChange(snap, sessionId)) return
      // A per-doc lastWriter filter here would risk dropping a genuine
      // remote change whose sibling doc happens to echo our own recent
      // write in the same snapshot. The snapshot carries the full set, so
      // the list is rebuilt from it rather than from the diff.
      {
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
        const createdDocuments = createdDocsFromDocs(snap.docs)
        if (createdDocuments.length > 0) safeUpdate({ createdDocuments })
      }
    })
  )

  active.unsubscribes.push(
    onSnapshot(mbRef, (snap) => {
      if (firstMbSnap) {
        firstMbSnap = false
        return
      }
      if (!hasRemoteChange(snap, sessionId)) return
      // Built from the snapshot's own full set, not a fresh re-read — same reasoning as todos. Never let an EMPTY cloud collection wipe the
      // on-screen list: a workspace still on the legacy misc.mbScrutiny blob has
      // an empty collection until its first push migrates it, and applying []
      // here would erase those records. The migrated records get pushed up
      // moments later; a genuine "deleted every MB" reflects on the next full pull.
      const mbScrutiny = mbFromDocs(snap.docs)
      if (mbScrutiny.length > 0) {
        if (active) active.pushedMbSig = new Map(mbScrutiny.map((v, i) => [v.id, JSON.stringify(mbFields(v, i))]))
        safeUpdate({ mbScrutiny })
      }
    })
  )

  let firstNoticeSnap = true
  active.unsubscribes.push(
    onSnapshot(noticesRef, (snap) => {
      if (firstNoticeSnap) {
        firstNoticeSnap = false
        return
      }
      if (!hasRemoteChange(snap, sessionId)) return
      // Same rule as MB Scrutiny: an EMPTY collection never wipes the on-screen list (a workspace still on the legacy blob has none until its first push migrates it).
      const pendingNotices = noticesFromDocs(snap.docs)
      if (pendingNotices.length > 0) {
        if (active) active.pushedNoticeSig = new Map(pendingNotices.map((v, i) => [v.id, JSON.stringify(noticeFields(v, i))]))
        safeUpdate({ pendingNotices })
      }
    })
  )
}

/**
 * Copies saved tender conditions into other users' cloud workspaces. Entries keep the office they were saved
 * for, so they appear for that user under that office (and stay out of every other one). A tender the user
 * already holds is left alone, and a user who has never opened the app (no workspace yet) is reported rather
 * than given a half-empty one. The write is a merge on the saved-tenders map only, so nothing else of theirs
 * is touched; their open app picks it up through its live listener.
 */
export async function shareTenderConditions(
  targetLoginIds: string[],
  entries: Record<string, TenderConditionsEntry>
): Promise<ShareResult[]> {
  await ensureAuth()
  const results: ShareResult[] = []
  const now = Date.now()
  for (const target of targetLoginIds) {
    const id = normalizeId(target)
    try {
      const miscRef = doc(db!, 'users', id, 'meta', 'misc')
      const snap = await getDoc(miscRef)
      if (!snap.exists()) {
        results.push({ loginId: target, added: 0, skipped: 0, error: 'has not used the app yet' })
        continue
      }
      const theirs = (snap.data().tenderConditions as Record<string, unknown> | undefined) ?? {}
      const toAdd: Record<string, TenderConditionsEntry> = {}
      let skipped = 0
      for (const [tid, entry] of Object.entries(cleanTenderConditionsForCloud(entries))) {
        if (tid in theirs) skipped++
        else {
          const { expiresAt: _drop, ...rest } = entry
          toAdd[tid] = { ...rest, sharedBy: active?.loginId ?? 'another user', savedAt: now }
        }
      }
      const added = Object.keys(toAdd).length
      if (added > 0) {
        await setDoc(
          miscRef,
          { tenderConditions: toAdd, updatedAt: serverTimestamp(), lastWriter: `share:${active?.sessionId ?? 'unknown'}` },
          { merge: true }
        )
      }
      results.push({ loginId: target, added, skipped })
    } catch (e) {
      console.error('shareTenderConditions failed for', target, e)
      results.push({ loginId: target, added: 0, skipped: 0, error: e instanceof Error ? e.message : String(e) })
    }
  }
  return results
}

/** Push the current workspace to the cloud, tagged as written by this device. Resolves true once everything was written, false if signed out or the push failed (autosave ignores this; the manual Sync button reports it). */
/** The shared meta/misc document's fields; `tenderConditions` is passed in because a push swaps removed IDs for deleteField(). */
/** Firestore rejects `undefined` anywhere in a document, and a record built with optional fields left unset (e.g. a pending Notice's itemNo/zone) carries them as explicit undefined keys — a JSON round-trip drops them. */
function plain<T>(v: T): T {
  return JSON.parse(JSON.stringify(v))
}

function buildMiscFields(state: PersistedState, tenderConditions: Record<string, unknown>, monitoringFormatByOffice: Record<string, unknown> = state.monitoringFormatByOffice ?? {}) {
  return {
        // Round-trip the schema version so one-time migrations run only once (see
        // pullRemoteState) instead of re-running — and compounding — every launch.
        version: state.version,
        resolution: plain(state.resolution ?? {}),
        todos: plain(state.todos ?? []),
        lastGoogleLink: plain(state.lastGoogleLink ?? null),
        worksListLinks: plain(state.worksListLinks ?? {}),
        // merge:true merges this map key by key, so a device with fewer saved contacts never wipes another's.
        contacts: plain(state.contacts ?? {}),
        qcParties: plain(state.qcParties ?? {}),
        monitoringFormatByOffice,
        monitoringFormatLinks: plain(state.monitoringFormatLinks ?? {}),
        tenderConditions,
        tenderReminders: plain(state.tenderReminders ?? []),
        pendingNotices: plain(state.pendingNotices ?? []),
        bidDocumentBatches: plain(state.bidDocumentBatches ?? []),
        mbScrutiny: plain(state.mbScrutiny ?? []),
        // Deliberately omitted (not written as null) when not yet chosen —
        // `merge: true` below then leaves whatever office is already stored in
        // the cloud untouched, matching PersistedState.office's own LWW
        // contract ("never let an empty office overwrite a good one already
        // stored"). Every other field above is unconditional/defaulted, so
        // switching this write to merge:true doesn't change their behavior.
        // Round-tripped through JSON so a zone-level office's undefined circle / circle number is dropped: Firestore rejects undefined values, which failed every misc push for such offices.
        ...(state.office ? { office: JSON.parse(JSON.stringify(state.office)) } : {})
    }
}

export async function pushState(state: PersistedState): Promise<boolean> {
  if (!active) {
    lastPushError = 'Not signed in to the cloud (the app started offline) — log out and log in again.'
    return false
  }
  lastPushError = ''
  const { loginId, sessionId } = active
  try {
    await ensureAuth()
    // Each office's Works List is its own document, so no single document can outgrow Firestore's 1 MiB cap
    // (the shared meta/tables one did). Only offices that changed since they were last read/written here are
    // sent, so an office another device edited is never overwritten by this device's stale copy.
    const pushedWorks = active?.pushedWorksSig
    for (const [officeK, list] of Object.entries(state.tablesByOffice ?? {})) {
      const json = JSON.stringify(list ?? [])
      const known = pushedWorks?.get(officeK)
      if (known === json) continue
      if (known === undefined && json === '[]') continue // never-used office the cloud doesn't hold — nothing to store
      await setDoc(doc(db!, 'users', loginId, 'worksByOffice', worksDocId(officeK)), {
        officeKey: officeK,
        tables: list ?? [],
        updatedAt: serverTimestamp(),
        lastWriter: sessionId
      })
      pushedWorks?.set(officeK, json)
    }
    // The shared document now keeps only the flat current-office copy (older app versions read it) and drops the
    // old all-offices map — rewritten WITHOUT merge so that map's weight is actually released. Runs after every
    // office above was written, so nothing is lost if a write failed (this line is never reached then).
    const tablesSig = JSON.stringify(state.tables ?? [])
    if (active && active.pushedTablesSig !== tablesSig) {
      await setDoc(doc(db!, 'users', loginId, 'meta', 'tables'), {
        tables: state.tables ?? [],
        perOffice: true,
        updatedAt: serverTimestamp(),
        lastWriter: sessionId
      })
      active.pushedTablesSig = tablesSig
    }
    // `merge: true` merges the saved-tenders map, so an ID deleted here must be removed with deleteField().
    const cleanedTenders = cleanTenderConditionsForCloud(state.tenderConditions)
    const tenderConditionsForPush: Record<string, unknown> = { ...cleanedTenders }
    const deletedTenders: string[] = []
    for (const id of knownCloudTenderIds) {
      if (!(id in cleanedTenders)) {
        tenderConditionsForPush[id] = deleteField()
        deletedTenders.push(id)
      }
    }
    // A Monitoring Format deleted here is removed from the cloud's merged map too, or the next pull would bring it back.
    const mfPlain = plain(state.monitoringFormatByOffice ?? {})
    const mfForPush: Record<string, unknown> = { ...mfPlain }
    const deletedMf: string[] = []
    for (const k of knownCloudMfKeys) {
      if (!(k in mfPlain)) {
        mfForPush[k] = deleteField()
        deletedMf.push(k)
      }
    }
    const miscFields = buildMiscFields(state, tenderConditionsForPush, mfForPush)
    // The deleteField() markers don't serialise, so the signature carries the plain saved-tenders map and the ids being deleted.
    const miscSig = JSON.stringify({ ...miscFields, tenderConditions: cleanedTenders, monitoringFormatByOffice: mfPlain, deletedTenders, deletedMf })
    if (active && active.pushedMiscSig !== miscSig) {
      await setDoc(
        doc(db!, 'users', loginId, 'meta', 'misc'),
        { ...miscFields, updatedAt: serverTimestamp(), lastWriter: sessionId },
        { merge: true }
      )
      active.pushedMiscSig = miscSig
    }
    knownCloudTenderIds = new Set(Object.keys(cleanedTenders))
    knownCloudMfKeys = new Set(Object.keys(mfPlain))

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
        lastPushError = e instanceof Error ? e.message : String(e)
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
      const fields = todoFields(t, i)
      const sig = JSON.stringify(fields)
      if (active?.pushedTodoSig.get(t.id) === sig) continue
      try {
        await setDoc(doc(db!, 'users', loginId, 'todos', t.id), {
          ...fields,
          updatedAt: serverTimestamp(),
          lastWriter: sessionId
        })
        active?.pushedTodoSig.set(t.id, sig)
      } catch (e) {
        console.error(`pushState: failed to sync todo "${t.text}"`, e)
        lastPushError = e instanceof Error ? e.message : String(e)
      }
    }
    if (active) {
      for (const oldId of active.knownTodoIds) {
        if (!currentTodoIds.has(oldId)) {
          await deleteDoc(doc(db!, 'users', loginId, 'todos', oldId))
          active.pushedTodoSig.delete(oldId)
        }
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
      const fields = mbFields(it, i)
      const sig = JSON.stringify(fields)
      if (active?.pushedMbSig.get(it.id) === sig) continue
      try {
        await setDoc(doc(db!, 'users', loginId, 'mbScrutiny', it.id), {
          ...fields,
          updatedAt: serverTimestamp(),
          lastWriter: sessionId
        })
        active?.pushedMbSig.set(it.id, sig)
      } catch (e) {
        console.error(`pushState: failed to sync MB record "${it.mbNo}"`, e)
        lastPushError = e instanceof Error ? e.message : String(e)
      }
    }
    if (active) {
      for (const oldId of active.knownMbIds) {
        if (!currentMbIds.has(oldId)) {
          await deleteDoc(doc(db!, 'users', loginId, 'mbScrutiny', oldId))
          active.pushedMbSig.delete(oldId)
        }
      }
      active.knownMbIds = currentMbIds
    }
    // Non-conclusion Notices: one Firestore doc per notice (mirrors MB Scrutiny), so a stale whole-blob
    // push can't wipe notices another device just generated. The legacy misc.pendingNotices stays dual-written.
    const orderedNotices = state.pendingNotices ?? []
    const currentNoticeIds = new Set(orderedNotices.map((it) => it.id))
    for (let i = 0; i < orderedNotices.length; i++) {
      const it = orderedNotices[i]
      const fields = noticeFields(it, i)
      const sig = JSON.stringify(fields)
      if (active?.pushedNoticeSig.get(it.id) === sig) continue
      try {
        await setDoc(doc(db!, 'users', loginId, 'pendingNotices', it.id), {
          ...fields,
          updatedAt: serverTimestamp(),
          lastWriter: sessionId
        })
        active?.pushedNoticeSig.set(it.id, sig)
      } catch (e) {
        console.error(`pushState: failed to sync notice "${it.nitNo}"`, e)
        lastPushError = e instanceof Error ? e.message : String(e)
      }
    }
    if (active) {
      for (const oldId of active.knownNoticeIds) {
        if (!currentNoticeIds.has(oldId)) {
          await deleteDoc(doc(db!, 'users', loginId, 'pendingNotices', oldId))
          active.pushedNoticeSig.delete(oldId)
        }
      }
      active.knownNoticeIds = currentNoticeIds
    }
    return lastPushError === ''
  } catch (e) {
    console.error('pushState failed', e)
    lastPushError = e instanceof Error ? e.message : String(e)
    return false
  }
}

/**
 * How many bytes the next push would upload: the same change detection as pushState (only what differs from what
 * was last read from / written to the cloud), so the UI can show it and auto-sync once it grows large. Deletions
 * count a token amount so they are not invisible. 0 means everything is already online (or not signed in).
 */
export function pendingBytes(state: PersistedState): number {
  if (!active) return 0
  const size = (v: string) => Buffer.byteLength(v, 'utf8')
  let total = 0
  for (const [officeK, list] of Object.entries(state.tablesByOffice ?? {})) {
    const json = JSON.stringify(list ?? [])
    const known = active.pushedWorksSig.get(officeK)
    if (known === json || (known === undefined && json === '[]')) continue
    total += size(json)
  }
  const tablesSig = JSON.stringify(state.tables ?? [])
  if (active.pushedTablesSig !== tablesSig) total += size(tablesSig)
  const cleanedTenders = cleanTenderConditionsForCloud(state.tenderConditions)
  const deletedTenders = [...knownCloudTenderIds].filter((id) => !(id in cleanedTenders))
  const mfPlain = state.monitoringFormatByOffice ?? {}
  const deletedMf = [...knownCloudMfKeys].filter((k) => !(k in mfPlain))
  const miscSig = JSON.stringify({ ...buildMiscFields(state, cleanedTenders), monitoringFormatByOffice: mfPlain, deletedTenders, deletedMf })
  if (active.pushedMiscSig !== miscSig) total += size(miscSig)
  const docs = state.createdDocuments ?? []
  docs.forEach((d, i) => {
    const sig = `${i}|${d.name}|${d.docx}`
    if (active!.pushedDocSig.get(d.id) !== sig) total += size(sig)
  })
  const docIds = new Set(docs.map((d) => d.id))
  total += 100 * [...active.knownDocIds].filter((id) => !docIds.has(id)).length
  const todos = state.todos ?? []
  todos.forEach((t, i) => {
    const sig = JSON.stringify(todoFields(t, i))
    if (active!.pushedTodoSig.get(t.id) !== sig) total += size(sig)
  })
  const todoIds = new Set(todos.map((t) => t.id))
  total += 100 * [...active.knownTodoIds].filter((id) => !todoIds.has(id)).length
  const mb = state.mbScrutiny ?? []
  mb.forEach((it, i) => {
    const sig = JSON.stringify(mbFields(it, i))
    if (active!.pushedMbSig.get(it.id) !== sig) total += size(sig)
  })
  const mbIds = new Set(mb.map((it) => it.id))
  total += 100 * [...active.knownMbIds].filter((id) => !mbIds.has(id)).length
  const notices = state.pendingNotices ?? []
  notices.forEach((it, i) => {
    const sig = JSON.stringify(noticeFields(it, i))
    if (active!.pushedNoticeSig.get(it.id) !== sig) total += size(sig)
  })
  const noticeIds = new Set(notices.map((it) => it.id))
  total += 100 * [...active.knownNoticeIds].filter((id) => !noticeIds.has(id)).length
  return total
}

import { useEffect, useMemo, useRef, useState, Fragment, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { api } from './ipc'
import { prefetchTenders, fetchTenders } from './tenderCache'
import {
  createWorksTable,
  applyWorksSchema,
  applyWorksSchemaWithMapping,
  migrateEcvContractToRupees,
  repairInflatedRupees,
  ECV_RUPEES_STATE_VERSION,
  WORKS_COLUMNS
} from './worksSchema'
import { autofillWorksRow, enforceZoneCircle, fillCircleNumber, splitCircleColumn } from './zoneCircleCheck'
import { closeOnBackdropMouseDown } from './overlayClose'
import { entriesOf, corporationByName } from './zoneCircleDirectory'
import { type Office, officeKey, isOfficeReady } from './office'
import { type ThemeId, getStoredTheme, setStoredTheme } from './theme'
import { matchPlaceholdersToColumns } from '@core/createDocument'
import { findWorksListErrors } from '@core/worksListAgent'
import { mergeTables } from '@core/merge'
import Sidebar, { type TabKey } from './components/Sidebar'
import QcPartiesEditor from './components/QcPartiesEditor'
import ExcelInline from './components/ExcelInline'
import CollisionPanel from './components/CollisionPanel'
import SearchTender from './components/SearchTender'
import Dashboard from './components/Dashboard'
import Skyline from './components/Skyline'
import ProfileMenu from './components/ProfileMenu'
import UpdateBanner from './components/UpdateBanner'
import TenderNoticeButton from './components/TenderNoticeButton'
import BidDocumentsPanel from './components/BidDocumentsPanel'
import SeBidDocumentTile from './components/SeBidDocumentTile'
import GoogleLinkImport from './components/GoogleLinkImport'
import WorksListL1Update from './components/WorksListL1Update'
import EstimateWorkspaceTab from './components/EstimateWorkspaceTab'
import GiveTechnicalSanctionTab from './components/GiveTechnicalSanctionTab'
import GiveIntimationTab from './components/GiveIntimationTab'
import EvaluationSheetTab from './components/EvaluationSheetTab'
import WorkOrderAgreementTab from './components/WorkOrderAgreementTab'
import IssueNoticesTab from './components/IssueNoticesTab'
import PrintDocumentTab from './components/PrintDocumentTab'
import ToolsTab from './components/ToolsTab'
import SettingsTab from './components/SettingsTab'
import TodoList from './components/TodoList'
import MbScrutinyList from './components/MbScrutinyList'
import MbMeasurementUploadTab from './components/MbMeasurementUploadTab'
import WhatsNew from './components/WhatsNew'
import { changesSince, CHANGELOG, type ChangelogEntry } from '@core/changelog'
import TenderReminders from './components/TenderReminders'
import {
  IconTable,
  IconDoc,
  IconPrint,
  IconSearch,
  IconCalendar,
  IconPlus,
  IconChecklist,
  IconClipboard,
  IconEye,
  IconRefresh,
  IconCheck,
  IconBell,
  IconDownload,
  IconUpload,
  IconTools,
  IconSettings,
  IconWarn,
  IconChevronRight
} from './components/Icons'
import type {
  ExcelTable,
  MergedDataset,
  CollisionResolution,
  TodoItem,
  MBScrutinyItem,
  TenderReminder,
  TenderReminderItem,
  CreatedDocument,
  BidDocumentBatch,
  BidDocumentWork,
  QcOfficeParties
} from '@core/types'
import type { CalendarData } from '@core/calendar'

// Older persisted reminders stored a single work (workName/tenderId/bidClosing)
// directly on the reminder instead of an `items` array — fold that into items.
function migrateTenderReminder(r: TenderReminder): TenderReminder {
  if (Array.isArray(r.items)) return r
  const legacy = r as TenderReminder & { workName?: string; tenderId?: string; bidClosing?: string }
  const item = { workName: legacy.workName, tenderId: legacy.tenderId, bidClosing: legacy.bidClosing }
  return { ...r, items: item.workName || item.tenderId || item.bidClosing ? [item] : [] }
}

// Older persisted MB Scrutiny entries stored remarks as a single string, then
// as plain point-wise strings, before each point got its own done checkbox —
// fold either legacy shape into the current { text, done } list.
function migrateMbScrutinyItem(it: MBScrutinyItem): MBScrutinyItem {
  const legacyRemarks = it.remarks as unknown
  if (typeof legacyRemarks === 'string') {
    return { ...it, remarks: legacyRemarks.trim() ? [{ text: legacyRemarks.trim(), done: false }] : [] }
  }
  if (Array.isArray(legacyRemarks) && legacyRemarks.some((r) => typeof r === 'string')) {
    return { ...it, remarks: (legacyRemarks as string[]).map((text) => ({ text, done: false })) }
  }
  return it
}

// Older persisted MB Scrutiny entries predate the running register number —
// assign one in stored (receipt) order to whichever entries are missing it.
function assignMissingSerialNos(items: MBScrutinyItem[]): MBScrutinyItem[] {
  let next = items.reduce((max, it) => Math.max(max, it.serialNo || 0), 0) + 1
  return items.map((it) => (it.serialNo ? it : { ...it, serialNo: next++ }))
}

let mbRestoreIdSeq = 0
function nextMbRestoreId(): string {
  mbRestoreIdSeq += 1
  return `mb-restore-${Date.now()}-${mbRestoreIdSeq}`
}

interface Props {
  onLogout: () => void
  /** The Head Draughtsman's chosen office (Corporation/Zone/Circle), selected in the sidebar — drives document prep and Works List validation. */
  office: Office
  onOfficeChange: (office: Office) => void
}

// The app version last shown in the "What's New" dialog on this machine —
// localStorage so it survives updates (but not a reinstall/cache clear).
const LAST_SEEN_VERSION_KEY = 'hda-last-seen-version'

/**
 * Keep a visited tab's content mounted (just hidden) instead of tearing it down
 * when the user navigates away, so any in-progress work (photo OCR, uploads,
 * long conversions) keeps running and the workspace's state is preserved when
 * they come back. A tab renders nothing until first visited; once visited it
 * stays mounted (display:none while another tab is active). Double-clicking a
 * sidebar item bumps a reset key upstream, which remounts these panes fresh —
 * that's the deliberate "close everything" gesture.
 */
function KeepAlive({ active, mounted, children }: { active: boolean; mounted: boolean; children: ReactNode }) {
  if (!active && !mounted) return null
  return <div className="tab-pane" style={{ display: active ? 'block' : 'none' }}>{children}</div>
}

export default function App({ onLogout, office, onOfficeChange }: Props) {
  // The office identity drives every Zone/Circle behaviour below. Kept as local
  // aliases so the rest of the app reads the same as when these came from the
  // login (Circle is optional — a zone-level Head Draughtsman picks just a Zone).
  const loginCorporation = office.corporation
  const loginZone = office.zone
  const loginCircle = office.circle
  const loginCircleNumber = office.circleNumber
  const officeEntries = entriesOf(loginCorporation)

  // "What's New" changes to announce on the first launch after an update (empty
  // until the version check below runs; a fresh install announces nothing).
  const [whatsNew, setWhatsNew] = useState<ChangelogEntry[]>([])

  // Issue Documents tile style (Settings → Themes) — a display preference for
  // this machine, not office-scoped data, so it lives in localStorage only
  // (see theme.ts).
  const [theme, setTheme] = useState<ThemeId>(() => getStoredTheme())
  function changeTheme(next: ThemeId) {
    setTheme(next)
    setStoredTheme(next)
  }
  // Windows theme and Dark mode both restyle the whole app (including
  // elements outside .shell, like Skyline/ProfileMenu, and portaled modals
  // on document.body), so they're scoped via body classes rather than a
  // container class.
  useEffect(() => {
    document.body.classList.toggle('theme-windows', theme === 'windows')
    document.body.classList.toggle('theme-dark', theme === 'dark')
  }, [theme])

  // Portal targets for the Agreement/Work Order and Intimation pages' own
  // "Download all documents" buttons — rendered in each page-head so they
  // sit next to the title (like Calendar's "Issue tender notice") instead
  // of in the body.
  const workOrderHeaderActionRef = useRef<HTMLDivElement>(null)
  const intimationHeaderActionRef = useRef<HTMLDivElement>(null)

  const [tab, setTab] = useState<TabKey>('dashboard')
  // Keep-alive bookkeeping: every tab the user has visited stays mounted (hidden)
  // so its work/state survives navigating away. `resetNonce` is part of the panes'
  // React key — bumping it (a sidebar double-click) remounts them all, the
  // explicit "terminate every workspace" gesture.
  const [mountedTabs, setMountedTabs] = useState<Set<TabKey>>(() => new Set<TabKey>(['dashboard']))
  const [resetNonce, setResetNonce] = useState(0)
  useEffect(() => {
    setMountedTabs((prev) => (prev.has(tab) ? prev : new Set(prev).add(tab)))
  }, [tab])
  // Show "What's New" once after an update: compare the running app version with
  // the one last seen on this machine (localStorage, so it persists across
  // updates but not across a reinstall). Newer → list what changed, then record
  // this version as seen only when the user dismisses it, so it reliably shows.
  // A fresh install (no version seen yet) records the version silently and shows
  // nothing — there's no prior release to announce an update from. In dev the
  // login screen is skipped, so this runs on launch rather than after login.
  useEffect(() => {
    let cancelled = false
    ;(async () => {
      const current = await api.getAppVersion()
      if (cancelled || !current) return
      const seen = localStorage.getItem(LAST_SEEN_VERSION_KEY)
      const entries = changesSince(seen, current)
      if (entries.length > 0) setWhatsNew(entries)
      else localStorage.setItem(LAST_SEEN_VERSION_KEY, current)
    })()
    return () => {
      cancelled = true
    }
  }, [])
  // Double-click a sidebar item: drop every other kept-alive workspace and remount
  // the chosen one from scratch (bumping resetNonce also resets it if it was the
  // current tab), so all in-progress workspace work is torn down.
  const resetToTab = (next: TabKey) => {
    setMountedTabs(new Set<TabKey>([next]))
    setResetNonce((n) => n + 1)
    setTab(next)
  }
  // Intimation and Agreement/Work Order keep their uploaded documents and
  // filled fields alive across tab switches (see KeepAlive above) — a "Clear"
  // button in each page's header lets the office wipe just THAT workspace's
  // in-progress upload before starting the next work, without losing every
  // other open tab's state the way the sidebar's whole-app reset would.
  // Bumping the key remounts only that one component from scratch.
  const [intimationInstanceKey, setIntimationInstanceKey] = useState(0)
  const [workOrderInstanceKey, setWorkOrderInstanceKey] = useState(0)
  const [issueNoticesInstanceKey, setIssueNoticesInstanceKey] = useState(0)
  // Sub-tabs within the Intimation workspace: the Intimation letter, or the
  // Bid Capacity Evaluation Sheet issued from a "View Bidders" PDF.
  const [intimationSubTab, setIntimationSubTab] = useState<'intimation' | 'evaluation'>('intimation')
  const [calendar, setCalendar] = useState<CalendarData | null>(null)

  const [tables, setTables] = useState<ExcelTable[]>([])
  // Each office's Works List, kept so switching offices restores the office's own
  // data instead of erasing it and re-importing. `tables` mirrors the current
  // office's entry. Key: officeKey(office).
  const [tablesByOffice, setTablesByOffice] = useState<Record<string, ExcelTable[]>>({})
  const currentOfficeKey = officeKey(office) ?? ''
  const [resolution, setResolution] = useState<CollisionResolution>({})

  const [todos, setTodos] = useState<TodoItem[]>([])
  const [mbScrutiny, setMbScrutiny] = useState<MBScrutinyItem[]>([])
  const [mbMeasurementOpen, setMbMeasurementOpen] = useState(false)
  const [lastGoogleLink, setLastGoogleLink] = useState<string | null>(null)
  const [refreshingWorks, setRefreshingWorks] = useState(false)
  const [refreshError, setRefreshError] = useState<string | null>(null)
  // After an "Update from L1", the matched Works List row indices to flash (with
  // a message under them). `token` remounts ExcelInline so it re-reads the just-
  // updated rows (it holds its own grid copy and doesn't otherwise re-sync).
  const [worksFlash, setWorksFlash] = useState<{ token: number; rows: number[]; message: string } | null>(null)
  // Set when the office (Zone/Circle) was just changed — each circle has its own
  // Works List, so we prompt the user to import that circle's database link. The
  // stored label is the new office's Circle (or Zone) shown in the notice.
  const [officeImportPrompt, setOfficeImportPrompt] = useState<string | null>(null)
  // Set while switching to an office that has a remembered Works List link
  // but wasn't loaded yet this session — the auto re-import this triggers
  // (see pendingOfficeImport's effect) is a real network fetch + parse of
  // however many hundred rows that office's sheet has, with nothing else on
  // screen to show for it otherwise: the office modal just closes and the
  // Works List page sits there looking stuck/slow with no explanation
  // (real report: "selecting circle office is taking time").
  const [officeSwitchBusy, setOfficeSwitchBusy] = useState<string | null>(null)
  // Works List link remembered per office (key: officeKey) — synced, so a
  // circle's database reloads on return and follows the user across systems.
  const [worksListLinks, setWorksListLinks] = useState<Record<string, string>>({})
  // 3rd/4th-party QC agencies remembered per office (key: officeKey) — entered
  // once on the Works List page and reused by the 3rd/4th-party QC letters.
  const [qcParties, setQcParties] = useState<Record<string, QcOfficeParties>>({})
  // A remembered link queued to auto-import once the office prop has updated to
  // the newly-selected office (so the import validates against the new office).
  const [pendingOfficeImport, setPendingOfficeImport] = useState<string | null>(null)
  // Outcome of the Google-link refresh (a plain re-import of the Works List).
  const [refreshSummary, setRefreshSummary] = useState<string | null>(null)
  const [tenderReminders, setTenderReminders] = useState<TenderReminder[]>([])
  const [refreshingReminderId, setRefreshingReminderId] = useState<string | null>(null)
  // The raw, unbaked templates (office placeholders like {{circle}} still intact)
  // — the source of truth that persists. `bakedDocuments` derives from these with
  // the *current* office stamped in, recomputed whenever the office changes so a
  // document never keeps a stale circle after the office is switched.
  const [createdDocuments, setCreatedDocuments] = useState<CreatedDocument[]>([])
  const [bakedDocuments, setBakedDocuments] = useState<CreatedDocument[]>([])
  // Highest bundled-default-documents version merged into this workspace. Set
  // by the main process's one-time injection on load; persisted here so the
  // injection never repeats (and a default the user deletes stays deleted).
  const [seededDocVersion, setSeededDocVersion] = useState(0)
  const [bidDocumentBatches, setBidDocumentBatches] = useState<BidDocumentBatch[]>([])
  // Set when the previously-saved Works List belongs to a different
  // Zone/Circle than the one logged in — that data is withheld (never
  // loaded into `tables`, so it never displays). Everything else (documents,
  // todos, reminders, ...) keeps saving normally; only the withheld tables
  // data itself is protected — see withheldTablesRef below.
  const [blockedWorksList, setBlockedWorksList] = useState<{ zone?: string; circle?: string } | null>(null)
  // The actual on-disk Works List data while it's withheld — substituted
  // back into every autosave in place of the (deliberately empty) `tables`
  // state, so the real data on disk is preserved untouched rather than
  // clobbered with `[]` the moment anything else changes.
  const withheldTablesRef = useRef<ExcelTable[] | null>(null)

  // A fresh import whose foreign-circle rows (see enforceZoneCircle) exceed
  // 20% of the sheet pauses here for the user to confirm it's really their
  // own office's works list before committing — a reorganisation can
  // legitimately leave a few old-circle rows behind, but a sheet that's
  // mostly someone else's circle is more likely the wrong file pasted in.
  const [circleMismatchConfirm, setCircleMismatchConfirm] = useState<{
    count: number
    total: number
    zone: string
    circle: string
  } | null>(null)
  const circleMismatchResolveRef = useRef<((proceed: boolean) => void) | null>(null)

  function confirmCircleMismatch(count: number, total: number, zone: string, circle: string): Promise<boolean> {
    return new Promise((resolve) => {
      circleMismatchResolveRef.current = resolve
      setCircleMismatchConfirm({ count, total, zone, circle })
    })
  }

  function resolveCircleMismatch(proceed: boolean) {
    circleMismatchResolveRef.current?.(proceed)
    circleMismatchResolveRef.current = null
    setCircleMismatchConfirm(null)
  }

  // Set for exactly the next autosave whenever state was just applied FROM a
  // remote sync — so that save persists locally but is NOT pushed back to the
  // cloud. Without this, receiving a remote change re-pushes it, the other
  // sessions re-push it in turn, and the resulting write storm lets a stale
  // copy overwrite (e.g.) a To Do task another session just added.
  const savedFromRemoteRef = useRef(false)

  // Holds the not-yet-fired debounced save below, so a logout (which unmounts
  // this component) can flush it immediately instead of losing it: the
  // effect's own cleanup just clearTimeout()s the pending save on every
  // dependency change (that's the debounce), which would silently drop the
  // latest edit if it happens to run as part of an unmount.
  const pendingSaveRef = useRef<null | (() => void)>(null)

  // The current Excel shown on the Data tab (single-workbook workflow).
  const currentTable = tables[0] ?? null

  // Debounced copy of currentTable for the error scan below — editing a
  // Works List cell replaces `tables` (a new array/object) on every single
  // keystroke, and without this the full Wincode/Tender-ID/ECV/EMD scan
  // (core/worksListAgent.ts, O(rows) but real regex/string work per row)
  // re-ran on every keystroke too, for however many hundred rows the list
  // has — a real contributor to "typing feels slow", especially on slower
  // hardware. The Errors button is a background health-check, not something
  // that needs to react within a keystroke, so it settling ~600ms after you
  // stop typing (rather than instantly on every character) changes nothing
  // about correctness, just when the count visibly updates.
  const [debouncedTable, setDebouncedTable] = useState(currentTable)
  useEffect(() => {
    const t = setTimeout(() => setDebouncedTable(currentTable), 600)
    return () => clearTimeout(t)
  }, [currentTable])

  // Every Works List error check (Wincode/Tender ID identity, ECV vs.
  // estimate, EMD 1%/1.5% consistency — see core/worksListAgent.ts) —
  // recomputed only when the (debounced) table itself changes, not on every
  // render, since it walks every row.
  const worksListErrors = useMemo(
    () => (debouncedTable ? findWorksListErrors(debouncedTable) : []),
    [debouncedTable]
  )
  const [showWincodeViolations, setShowWincodeViolations] = useState(false)

  // Global dataset (all Excels merged) — used for collision detection on the
  // Works List tab.
  const dataset: MergedDataset | null = useMemo(
    () => (tables.length === 0 ? null : mergeTables(tables, resolution)),
    [tables, resolution]
  )

  // Bake Zone/Circle/CNO straight into a document's placeholders when logged
  // in at Circle level — every work under that login belongs to the same
  // Zone/Circle/CNO, so there's no ambiguity to leave for print time. A
  // Zone-level login (loginCircle unset) spans many circles' works, so those
  // placeholders must stay dynamic instead, resolved per-row as before.
  async function bakeLoginPlaceholders(docs: CreatedDocument[]): Promise<CreatedDocument[]> {
    if (!loginZone || !loginCircle) return docs
    const values: Record<string, string> = { zone: loginZone, circle: loginCircle }
    if (loginCircleNumber) values.cno = loginCircleNumber
    // Corporation is part of the office too: {{Corporation}} (abbreviation) and
    // {{Corporation Full Name}} (uppercase title) so a document's letterhead
    // follows the chosen corporation, not a hard-coded one.
    if (loginCorporation) {
      values.corporation = loginCorporation
      const full = corporationByName(loginCorporation)?.fullName
      if (full) values['corporation full name'] = full.toUpperCase()
    }
    return Promise.all(
      docs.map(async (d) => ({ ...d, docx: await api.bakeFixedPlaceholdersInDocument(d.docx, values) }))
    )
  }

  // Re-bake the raw templates with the current office whenever either changes, so
  // every document (Issue Documents tab and the Tools blank forms) always shows
  // the office in use — switching Circle updates them instead of keeping the old
  // one baked in. A zone-only office leaves the placeholders for print-time.
  useEffect(() => {
    let cancelled = false
    void (async () => {
      const baked = await bakeLoginPlaceholders(createdDocuments)
      if (!cancelled) setBakedDocuments(baked)
    })()
    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [createdDocuments, loginZone, loginCircle, loginCircleNumber, loginCorporation])

  // Reorder acts on the raw list: the tabs show `bakedDocuments`, so a drag hands
  // back the baked docs in the new order — map that id order onto the raw list.
  function reorderDocuments(reordered: CreatedDocument[]) {
    const order = reordered.map((d) => d.id)
    setCreatedDocuments((prev) => {
      const byId = new Map(prev.map((d) => [d.id, d]))
      const next = order.map((id) => byId.get(id)).filter((d): d is CreatedDocument => !!d)
      // Keep any raw doc that wasn't in the reordered (baked) view, just in case.
      for (const d of prev) if (!order.includes(d.id)) next.push(d)
      return next
    })
  }

  // ── Persistence: load the saved workspace once on startup ──────────
  const [hydrated, setHydrated] = useState(false)
  useEffect(() => {
    // Warm the Search Tender default view in the background so the tab is
    // instant whenever the user opens it.
    prefetchTenders()
    let cancelled = false
    ;(async () => {
      try {
        const s = await api.loadState()
        if (!cancelled && s) {
          // Force every existing works database onto the standard column schema
          // (columns come from the APP Excel). Any other columns are dropped.
          // A Works List saved before ECV/Contract Amount moved to rupee
          // storage is converted once here (Lakhs -> rupees), keyed on the
          // persisted schema version so it never double-converts.
          const needsEcvRupeeMigration = (s.version ?? 1) < ECV_RUPEES_STATE_VERSION
          const migrate = (ts: ExcelTable[]): ExcelTable[] =>
            ts
              .map(applyWorksSchema)
              .map((t) => (needsEcvRupeeMigration ? migrateEcvContractToRupees(t) : t))
              // Always heal any ECV/Contract Amount left absurdly inflated by an
              // earlier over-migration (idempotent — see repairInflatedRupees).
              .map(repairInflatedRupees)

          // Restore each office's own Works List. Legacy states stored only a
          // single top-level `tables` (the last office in use) — migrate it into
          // this office's slot so switching away and back no longer loses it.
          const byOffice: Record<string, ExcelTable[]> = {}
          for (const [k, v] of Object.entries(s.tablesByOffice ?? {})) byOffice[k] = migrate(v)
          const loadedTables = byOffice[currentOfficeKey] ?? migrate(s.tables ?? [])
          if (!byOffice[currentOfficeKey] && loadedTables.length > 0) byOffice[currentOfficeKey] = loadedTables
          setTablesByOffice(byOffice)

          // A previously-saved Works List belonging to a different Zone/Circle
          // than the one logged in now must not be shown — same rule as a
          // fresh import (see importFromGoogleLink), applied to what's already
          // on disk too.
          const first = loadedTables[0] ? splitCircleColumn(loadedTables[0]) : undefined
          if (loginZone && loginCircle && first) {
            const { table: checked, mismatches } = enforceZoneCircle(first, loginZone, loginCircle, officeEntries)
            // A circle reorganisation (e.g. Moosapet's wards split into
            // Kukatpally) means some of the office's own works can legitimately
            // still carry their old circle's name — those foreign-tagged rows
            // are kept as-is, not rejected. Only block when NOT ONE row belongs
            // to the office's own Zone/Circle (the sheet is plainly the wrong
            // one), matching the fresh-import check below.
            if (mismatches.length > 0 && mismatches.length === checked.rows.length) {
              const m = mismatches[0]
              withheldTablesRef.current = loadedTables
              setBlockedWorksList({ zone: m.foundZone, circle: m.foundCircle })
            } else {
              setTables([fillCircleNumber(checked, loginCircleNumber), ...loadedTables.slice(1)])
            }
          } else {
            setTables(loadedTables)
          }
          setResolution(s.resolution ?? {})
          setTodos(s.todos ?? [])
          setMbScrutiny(assignMissingSerialNos((s.mbScrutiny ?? []).map(migrateMbScrutinyItem)))
          setLastGoogleLink(s.lastGoogleLink ?? null)
          setWorksListLinks(s.worksListLinks ?? {})
          setQcParties(s.qcParties ?? {})
          setTenderReminders((s.tenderReminders ?? []).map(migrateTenderReminder))
          setCreatedDocuments(s.createdDocuments ?? [])
          setSeededDocVersion(s.seededDocVersion ?? 0)
          setBidDocumentBatches(s.bidDocumentBatches ?? [])
        }
      } finally {
        if (!cancelled) setHydrated(true)
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  // ── Sync: apply changes pushed live from the other signed-in device ─
  // Only meaningful once the initial load has finished, so a remote update
  // can't race the startup load and get clobbered by it.
  useEffect(() => {
    if (!hydrated) return
    return api.onRemoteStateUpdate(async (partial) => {
      // This state came from the cloud — mark the autosave it triggers as
      // local-only so we don't echo it straight back (see savedFromRemoteRef).
      savedFromRemoteRef.current = true
      // Merge the offices the other session changed into our per-office store,
      // keeping every office it didn't touch — so a session working on another
      // office never wipes ours.
      if (partial.tablesByOffice) {
        const incoming: Record<string, ExcelTable[]> = {}
        for (const [k, v] of Object.entries(partial.tablesByOffice)) incoming[k] = v.map(applyWorksSchema)
        setTablesByOffice((prev) => ({ ...prev, ...incoming }))
        // Reflect it on screen only when it's this office's own data.
        const mine = incoming[currentOfficeKey]
        if (mine) {
          if (blockedWorksList) withheldTablesRef.current = mine
          else setTables(mine)
        }
      } else if (partial.tables) {
        // Legacy remote (older client) sends only the flat `tables`.
        const loadedTables = partial.tables.map(applyWorksSchema)
        // Keep the same Zone/Circle withholding behaviour as the initial
        // load: don't surface a Works List for a different office.
        if (blockedWorksList) {
          withheldTablesRef.current = loadedTables
        } else {
          setTables(loadedTables)
        }
      }
      if (partial.resolution) setResolution(partial.resolution)
      if (partial.todos) setTodos(partial.todos)
      if (partial.mbScrutiny)
        setMbScrutiny(assignMissingSerialNos(partial.mbScrutiny.map(migrateMbScrutinyItem)))
      if (partial.lastGoogleLink !== undefined) setLastGoogleLink(partial.lastGoogleLink ?? null)
      if (partial.worksListLinks) setWorksListLinks(partial.worksListLinks)
      if (partial.qcParties) setQcParties(partial.qcParties)
      if (partial.tenderReminders) setTenderReminders(partial.tenderReminders.map(migrateTenderReminder))
      if (partial.createdDocuments) setCreatedDocuments(partial.createdDocuments)
      if (partial.seededDocVersion !== undefined) setSeededDocVersion(partial.seededDocVersion)
      if (partial.bidDocumentBatches) setBidDocumentBatches(partial.bidDocumentBatches)
    })
  }, [hydrated, blockedWorksList, currentOfficeKey])

  // Keep the current office's slot in the per-office store in step with what's
  // on screen, so switching away and back restores exactly this. Skipped while
  // the list is withheld for a Zone/Circle mismatch (tables is empty then).
  useEffect(() => {
    if (!hydrated || !currentOfficeKey || blockedWorksList) return
    setTablesByOffice((prev) => (prev[currentOfficeKey] === tables ? prev : { ...prev, [currentOfficeKey]: tables }))
  }, [tables, currentOfficeKey, hydrated, blockedWorksList])

  // ── Persistence: save the workspace whenever it changes (debounced) ─
  // While a Works List is withheld for a Zone/Circle mismatch
  // (blockedWorksList), `tables` is deliberately empty in-memory — so we
  // substitute the real on-disk data (withheldTablesRef) back in for that
  // one field. Everything else (documents, todos, reminders, ...) always
  // saves from live state, so it's never paused or lost.
  useEffect(() => {
    if (!hydrated) return
    // If this state change was applied from a remote sync, save it to disk but
    // don't push it back to the cloud (which already has it) — otherwise the
    // echo storms across concurrent sessions. Consume the flag now, at effect
    // time: a genuinely local change landing before the debounce fires re-runs
    // this effect with the flag cleared, so it pushes normally.
    const skipCloud = savedFromRemoteRef.current
    savedFromRemoteRef.current = false
    const save = () => {
      pendingSaveRef.current = null
      const currentTables = blockedWorksList ? (withheldTablesRef.current ?? tables) : tables
      // Persist every office's data, with this office's slot reflecting the
      // (possibly withheld) real data. currentOfficeKey routes the cloud sync so
      // only this office's entry is updated — other offices stay put.
      const byOffice = currentOfficeKey ? { ...tablesByOffice, [currentOfficeKey]: currentTables } : tablesByOffice
      api.saveState(
        {
          version: ECV_RUPEES_STATE_VERSION,
          tables: currentTables,
          tablesByOffice: byOffice,
          currentOfficeKey: currentOfficeKey || undefined,
          // Persist the chosen office so it's restored on the next login / a
          // fresh machine (AuthGate reads it back when localStorage has none).
          // Only save a real selection — never let an empty office overwrite a
          // good one already stored (LWW) for other sessions.
          office: isOfficeReady(office) ? office : undefined,
          resolution,
          todos,
          lastGoogleLink: lastGoogleLink ?? undefined,
          worksListLinks,
          qcParties,
          tenderReminders,
          createdDocuments,
          seededDocVersion,
          bidDocumentBatches,
          mbScrutiny
        },
        skipCloud
      )
    }
    pendingSaveRef.current = save
    const handle = setTimeout(save, 400)
    return () => {
      clearTimeout(handle)
      if (pendingSaveRef.current === save) pendingSaveRef.current = null
    }
  }, [
    hydrated,
    blockedWorksList,
    tables,
    tablesByOffice,
    currentOfficeKey,
    office,
    resolution,
    todos,
    lastGoogleLink,
    worksListLinks,
    qcParties,
    tenderReminders,
    createdDocuments,
    seededDocVersion,
    bidDocumentBatches,
    mbScrutiny
  ])

  // To Do and MB Scrutiny are per-office, like the Works List: show only the
  // current office's items, and stamp each new item with the office it's added
  // under. The full (all-office) arrays stay in `todos`/`mbScrutiny` for sync;
  // these adapters slice/merge just the current office's slice for the tabs.
  const officeTodos = todos.filter((t) => (t.officeKey ?? '') === currentOfficeKey)
  const officeMb = mbScrutiny.filter((it) => (it.officeKey ?? '') === currentOfficeKey)
  // Bid document batches, same per-office scoping — otherwise a batch issued
  // for one office keeps showing under every other office switched to.
  const officeBidBatches = bidDocumentBatches.filter((b) => (b.officeKey ?? '') === currentOfficeKey)
  // Short label of the office these per-office lists are scoped to, for the page header.
  const officeLabel = office.circle || office.zone || ''
  function setOfficeTodos(next: TodoItem[]) {
    const stamped = currentOfficeKey
      ? next.map((t) => (t.officeKey ? t : { ...t, officeKey: currentOfficeKey }))
      : next
    setTodos([...todos.filter((t) => (t.officeKey ?? '') !== currentOfficeKey), ...stamped])
  }
  function setOfficeMb(next: MBScrutinyItem[]) {
    const stamped = currentOfficeKey
      ? next.map((it) => (it.officeKey ? it : { ...it, officeKey: currentOfficeKey }))
      : next
    setMbScrutiny([...mbScrutiny.filter((it) => (it.officeKey ?? '') !== currentOfficeKey), ...stamped])
  }

  // One-time adoption: To Do / MB records created before per-office scoping have
  // no officeKey and would otherwise show under every office. On first load with
  // an office in view, assign them to it so they settle into one office's list.
  useEffect(() => {
    if (!hydrated || !currentOfficeKey) return
    setTodos((prev) =>
      prev.some((t) => !t.officeKey) ? prev.map((t) => (t.officeKey ? t : { ...t, officeKey: currentOfficeKey })) : prev
    )
    setMbScrutiny((prev) =>
      prev.some((it) => !it.officeKey)
        ? prev.map((it) => (it.officeKey ? it : { ...it, officeKey: currentOfficeKey }))
        : prev
    )
    setBidDocumentBatches((prev) =>
      prev.some((b) => !b.officeKey) ? prev.map((b) => (b.officeKey ? b : { ...b, officeKey: currentOfficeKey })) : prev
    )
  }, [hydrated, currentOfficeKey])

  const collisions = dataset?.collisions ?? []
  const unresolved = collisions.filter((c) => !resolution[c.column])

  function resolveCollision(column: string, source: string) {
    setResolution((r) => ({ ...r, [column]: source }))
  }

  // Rename the works database (display name only).
  function renameTable(name: string) {
    if (!currentTable) return
    updateTable({ ...currentTable, name })
  }

  // Save the current Works List as an .xlsx (native save dialog in main).
  // Download the Works List as .xlsx. When it was imported from a Google sheet,
  // the file reproduces exactly that sheet's columns in the sheet's own order
  // (currentTable.sourceHeaders) — not the app's internal schema order, and
  // without the app-only columns the sheet didn't have — so the download mirrors
  // the source sheet. The built-in database (no sourceHeaders) exports as-is.
  async function exportWorksList() {
    if (!currentTable) return
    const headers = currentTable.sourceHeaders?.length ? currentTable.sourceHeaders : currentTable.headers
    await api.exportTable({ ...currentTable, headers }, currentTable.name || 'Works List')
  }

  // Flatten the MB Scrutiny register into a table and save it as an .xlsx.
  async function exportMbScrutiny() {
    if (officeMb.length === 0) return
    const fmt = (iso?: string) => {
      if (!iso) return ''
      const [y, m, d] = iso.split('-')
      return `${d}.${m}.${y}`
    }
    const headers = [
      'S.No',
      'MB No.',
      'Agency name',
      'Received date',
      'Target date',
      'Status',
      'Scrutiny completed date',
      'Remarks / objections'
    ]
    const rows = [...officeMb]
      .sort((a, b) => (a.serialNo || 0) - (b.serialNo || 0))
      .map((it) => ({
        'S.No': String(it.serialNo ?? ''),
        'MB No.': it.mbNo,
        'Agency name': it.agencyName,
        'Received date': fmt(it.receivedDate),
        'Target date': fmt(it.targetDate),
        Status: it.done ? 'Completed' : 'Pending',
        'Scrutiny completed date': fmt(it.completedDate),
        'Remarks / objections': (it.remarks ?? [])
          .map((r) => `${r.done ? '[x]' : '[ ]'} ${r.text}`)
          .join('\n')
      }))
    await api.exportTable({ id: 'mb-scrutiny', name: 'MB Scrutiny', path: '', headers, rows }, 'MB Scrutiny list')
  }

  // Rebuild MB Scrutiny records from a previously-downloaded .xlsx (same
  // column layout as exportMbScrutiny above) — a safety net for records lost
  // to a sync hiccup, an accidental delete, or a fresh machine. Records
  // already present (matched by MB No. + Agency + Received date) are left
  // alone, so restoring the same file twice is harmless.
  async function restoreMbScrutiny() {
    const tables = await api.pickExcels()
    if (tables.length === 0) return
    const parseDMY = (s: string): string | undefined => {
      const m = s.trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/)
      if (!m) return undefined
      const [, d, mo, y] = m
      return `${y}-${mo.padStart(2, '0')}-${d.padStart(2, '0')}`
    }
    const restored: MBScrutinyItem[] = tables
      .flatMap((t) => t.rows)
      .map((r) => {
        const mbNo = (r['MB No.'] ?? '').trim()
        const agencyName = (r['Agency name'] ?? '').trim()
        const receivedDate = parseDMY(r['Received date'] ?? '')
        if (!mbNo || !agencyName || !receivedDate) return null
        const serialNo = Number(r['S.No'])
        const remarks = (r['Remarks / objections'] ?? '')
          .split('\n')
          .map((line) => line.trim())
          .filter(Boolean)
          .map((line) => ({ text: line.replace(/^\[[x ]\]\s*/, '').trim(), done: line.startsWith('[x]') }))
        const item: MBScrutinyItem = {
          id: nextMbRestoreId(),
          serialNo: Number.isFinite(serialNo) && serialNo > 0 ? serialNo : 0,
          mbNo,
          agencyName,
          receivedDate,
          targetDate: parseDMY(r['Target date'] ?? ''),
          done: (r['Status'] ?? '').trim().toLowerCase() === 'completed',
          completedDate: parseDMY(r['Scrutiny completed date'] ?? ''),
          remarks: remarks.length > 0 ? remarks : undefined
        }
        return item
      })
      .filter((it): it is MBScrutinyItem => it !== null)
    if (restored.length === 0) return
    const existingKey = (it: MBScrutinyItem) => `${it.mbNo}|${it.agencyName}|${it.receivedDate}`
    const existingKeys = new Set(officeMb.map(existingKey))
    const toAdd: MBScrutinyItem[] = []
    const seen = new Set<string>()
    for (const it of restored) {
      const key = existingKey(it)
      if (existingKeys.has(key) || seen.has(key)) continue
      seen.add(key)
      toAdd.push(it)
    }
    if (toAdd.length === 0) return
    setOfficeMb(assignMissingSerialNos([...officeMb, ...toAdd]))
  }

  // Create the built-in works database (standard APP columns, one blank row).
  function createDatabase() {
    setTables((prev) => (prev.length > 0 ? prev : [createWorksTable()]))
    setBlockedWorksList(null)
  }

  // Download a pasted Google Sheets / Drive link and replace the works
  // database with it: the old columns and rows are dropped entirely, and a
  // fresh table is built from the link's rows, forced onto the standard
  // schema. Each standard column is filled from whichever imported column
  // means the same thing (semantic/keyword match, not exact name equality —
  // see applyWorksSchemaWithMapping) since a source sheet very often uses
  // different column names (e.g. "Estimate Amount" instead of the app's
  // "Amount of estimate") that an exact-name match would otherwise drop.
  // A sheet with only a header row (no works filled in yet) still updates
  // the database's columns to match — it isn't rejected just for having no
  // rows yet, since the point of importing it is often exactly to set up
  // the right columns before adding works.
  async function importFromGoogleLink(url: string): Promise<{ added: number; table: ExcelTable }> {
    const imported = await api.importFromLink(url)
    const rows = imported.rows.filter((row) => Object.values(row).some((v) => (v ?? '').trim() !== ''))

    let embeddings: { labelVectors: number[][]; columnVectors: number[][] } | undefined
    try {
      const [labelVectors, columnVectors] = await Promise.all([
        api.embedTexts(WORKS_COLUMNS),
        api.embedTexts(imported.headers)
      ])
      embeddings = { labelVectors, columnVectors }
    } catch {
      // Neural matching unavailable — matchPlaceholdersToColumns falls back
      // to plain token overlap automatically when no embeddings are passed.
      embeddings = undefined
    }
    // uniqueColumns: each imported column feeds at most one Works List column,
    // so e.g. "Name of the Agency" and "Address of the agency" can't both pull
    // from a single "Agency Details" column.
    const mapping = matchPlaceholdersToColumns(WORKS_COLUMNS, imported.headers, embeddings, { uniqueColumns: true })

    // Split any combined "57-Gajularamaram"/"Gajularamaram 57" Circle cell into
    // a bare Circle plus a Circle number before matching against the login.
    const normalized = splitCircleColumn(
      applyWorksSchemaWithMapping(imported.headers, rows, mapping, {
        id: `works-${Date.now()}`,
        name: 'Works database',
        path: ''
      })
    )

    // A works list is accepted as long as at least one row belongs to the
    // logged-in Head Draughtsman's own Zone/Circle. A row explicitly tagged
    // with a different Zone/Circle (or whose "Name of the work" names a
    // different one) is kept as-is rather than dropped or overwritten — CMC's
    // circle reorganisation split some circles' wards into others (e.g.
    // Moosapet's wards into Kukatpally), and since a work's Wincode/name never
    // changes, a genuinely-this-office's works list can legitimately carry a
    // handful of rows still tagged with their old, now-merged-away circle.
    // Only when NOT ONE row matches the login's own Zone/Circle is the sheet
    // rejected outright, as the plainly wrong file for this office.
    let result: ExcelTable
    if (loginZone && loginCircle) {
      const { table: checked, mismatches } = enforceZoneCircle(normalized, loginZone, loginCircle, officeEntries)
      if (mismatches.length > 0 && mismatches.length === checked.rows.length) {
        throw new Error(
          `This works list doesn't contain a single work from your office's Zone/Circle (${loginZone} / ${loginCircle}) — ` +
            `check the "Name of the work" / Zone / Circle columns, or that this is the right sheet for this office.`
        )
      }
      // A handful of foreign-circle rows is expected after a reorganisation,
      // but more than a fifth of the whole sheet is worth a second look
      // before committing — ask, rather than silently importing or rejecting.
      if (mismatches.length > 0 && mismatches.length / checked.rows.length > 0.2) {
        const proceed = await confirmCircleMismatch(mismatches.length, checked.rows.length, loginZone, loginCircle)
        if (!proceed) {
          throw new Error("Import cancelled — the works list wasn't changed.")
        }
      }
      result = fillCircleNumber(checked, loginCircleNumber)
    } else {
      result = normalized
    }
    setTables([result])
    setBlockedWorksList(null)
    setLastGoogleLink(url)
    setOfficeImportPrompt(null)
    // Remember this link as the current office's own Works List database, so it
    // reloads on return to this circle and syncs to the user's other systems.
    const key = officeKey(office)
    if (key) setWorksListLinks((prev) => (prev[key] === url ? prev : { ...prev, [key]: url }))

    return { added: rows.length, table: result }
  }

  // Re-pull the Works List from whichever Google link it was last filled from,
  // so the user doesn't have to re-paste the URL to get fresh data. The sheet is
  // the single source of truth, each work identified by its own Wincode. (This
  // used to also search the tender portal by Circle and stamp ECV/Tender ID/
  // Notice No onto rows by matching work name — but the portal carries no
  // Wincode, so it could only match by near-identical work name, which collapsed
  // many distinct works onto one tender and overwrote their ECVs. That sync is
  // gone; award data comes from the sheet itself or the "Update from L1" button.)
  async function refreshWorksList() {
    if (!lastGoogleLink) return
    setRefreshingWorks(true)
    setRefreshError(null)
    setRefreshSummary(null)
    try {
      const { added } = await importFromGoogleLink(lastGoogleLink)
      setRefreshSummary(`Refreshed from the Google link — ${added} work${added === 1 ? '' : 's'} imported.`)
    } catch (e) {
      setRefreshError(e instanceof Error ? e.message : String(e))
    } finally {
      setRefreshingWorks(false)
    }
  }

  // Search the live e-procurement portal for a NIT number and fold every
  // matching row (a single NIT can cover several works/lots) into the
  // given reminder's items list.
  async function lookupTenderReminder(id: string, nitNo: string) {
    try {
      const res = await fetchTenders({ start: 0, length: 25, type: 'current', search: nitNo }, true)
      const matches = res.data.filter((row) => (row[2] || '').trim() === nitNo)
      const items: TenderReminderItem[] = matches.map((row) => ({
        workName: row[4] || undefined,
        tenderId: row[1] || undefined,
        bidClosing: row[7] || undefined
      }))
      setTenderReminders((prev) =>
        prev.map((r) =>
          r.id !== id ? r : items.length > 0 ? { ...r, status: 'found', items } : { ...r, status: 'not-found' }
        )
      )
    } catch {
      setTenderReminders((prev) => prev.map((r) => (r.id === id ? { ...r, status: 'not-found' } : r)))
    }
  }

  // Called once a tender notice is generated: save the NIT number as a
  // reminder and immediately look it up on the Search Tender portal.
  async function addTenderReminder(nitNo: string) {
    const id = `reminder-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
    const d = new Date()
    const createdDate = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
    setTenderReminders((prev) => [...prev, { id, nitNo, items: [], status: 'pending', createdDate }])
    await lookupTenderReminder(id, nitNo)
  }

  function removeTenderReminder(id: string) {
    setTenderReminders((prev) => prev.filter((r) => r.id !== id))
  }

  function addBidBatch(batch: BidDocumentBatch) {
    setBidDocumentBatches((prev) => [...prev, currentOfficeKey ? { ...batch, officeKey: currentOfficeKey } : batch])
  }

  function removeBidBatch(id: string) {
    setBidDocumentBatches((prev) => prev.filter((b) => b.id !== id))
  }

  /** Edits one work's SE-only manual fields (Item No, TS No/Date, Admin Sanction) in place — the EE fields (Name/Amount/ECV/Zone/Circle) always come from the Works List and are never user-edited here. */
  function updateBidWork(batchId: string, serial: number, patch: Partial<BidDocumentWork>) {
    setBidDocumentBatches((prev) =>
      prev.map((b) =>
        b.id !== batchId ? b : { ...b, works: b.works.map((w) => (w.serial !== serial ? w : { ...w, ...patch })) }
      )
    )
  }

  // Add (or fold into, if a reminder with the same NIT number already
  // exists) a reminder straight from a Search Tender result row — no extra
  // lookup needed, the data's already at hand. A single NIT number can
  // cover multiple works, so each distinct Tender ID becomes its own item.
  function addTenderReminderFromRow(row: string[]) {
    const nitNo = (row[2] || '').trim()
    if (!nitNo) return
    const item: TenderReminderItem = {
      workName: row[4] || undefined,
      tenderId: row[1] || undefined,
      bidClosing: row[7] || undefined
    }
    setTenderReminders((prev) => {
      const idx = prev.findIndex((r) => r.nitNo === nitNo)
      if (idx !== -1) {
        const existing = prev[idx]
        const itemIdx = existing.items.findIndex((it) => item.tenderId && it.tenderId === item.tenderId)
        const items =
          itemIdx !== -1
            ? existing.items.map((it, i) => (i === itemIdx ? item : it))
            : [...existing.items, item]
        const next = [...prev]
        next[idx] = { ...existing, status: 'found', items }
        return next
      }
      const id = `reminder-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
      const d = new Date()
      const createdDate = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
      return [...prev, { id, nitNo, items: [item], status: 'found', createdDate }]
    })
  }

  async function refreshTenderReminder(id: string) {
    const r = tenderReminders.find((x) => x.id === id)
    if (!r) return
    setRefreshingReminderId(id)
    setTenderReminders((prev) => prev.map((x) => (x.id === id ? { ...x, status: 'pending' } : x)))
    await lookupTenderReminder(id, r.nitNo)
    setRefreshingReminderId(null)
  }

  function updateTable(updated: ExcelTable) {
    setTables((prev) => prev.map((t) => (t.id === updated.id ? updated : t)))
  }

  // Live auto-fill for the Works List grid: as a work name (or Circle) is typed,
  // derive that row's blank Zone / Circle / Circle number from it.
  const autofillWorksRowForLogin = (row: Record<string, string>) =>
    autofillWorksRow(row, {
      corporation: loginCorporation,
      zone: loginZone,
      circle: loginCircle,
      circleNumber: loginCircleNumber
    })

  // Changing the office changes which Works List applies — each circle (or a
  // zonal office's zone) has its own database. We save the office we're leaving
  // and RESTORE the office we're switching to from its own saved data, so its
  // database (including local edits) is never erased. Only an office we've never
  // loaded before falls back to importing its link / prompting for one.
  function changeOffice(next: Office) {
    const changed =
      next.corporation !== office.corporation || next.zone !== office.zone || next.circle !== office.circle
    const prevKey = currentOfficeKey
    onOfficeChange(next)
    if (changed && next.corporation && next.zone) {
      const nextKey = officeKey(next) ?? ''
      // Preserve the office we're leaving (unless it's withheld, in which case
      // its real data already lives in withheldTablesRef / tablesByOffice).
      if (prevKey && !blockedWorksList) {
        setTablesByOffice((prev) => ({ ...prev, [prevKey]: tables }))
      }
      setBlockedWorksList(null)
      withheldTablesRef.current = null
      // A "Refreshed from the Google link — N works imported." banner (or its
      // error counterpart) is a status for the office it just ran under —
      // otherwise it keeps showing (misleadingly claiming a fresh import, or
      // flagging an error that doesn't apply) after switching to an office
      // that was never actually refreshed this session.
      setRefreshSummary(null)
      setRefreshError(null)
      setTab('data')

      const stored = tablesByOffice[nextKey]
      if (stored && stored.length > 0) {
        // We already have this office's database — restore it, don't re-import.
        setTables(stored)
        setLastGoogleLink(worksListLinks[nextKey] ?? null)
        setOfficeImportPrompt(null)
        setPendingOfficeImport(null)
        return
      }

      // First time for this office in this session — clear and import/prompt.
      setTables([])
      setLastGoogleLink(null)
      const savedLink = worksListLinks[nextKey]
      if (savedLink) {
        // This office has a remembered link — reload its database automatically.
        // Deferred to an effect so the import validates against the now-updated
        // office (see below).
        setOfficeImportPrompt(null)
        setPendingOfficeImport(savedLink)
      } else {
        // No data and no link yet — ask the user for its link.
        setPendingOfficeImport(null)
        setOfficeImportPrompt(next.circle || next.zone)
      }
    }
  }

  // Run a queued office switch's remembered-link import once the office prop has
  // updated to the new office, so importFromGoogleLink validates against it. If
  // the remembered link fails (revoked/renamed), fall back to prompting. The ref
  // dedupes StrictMode's double-invoke so it imports once.
  const officeImportRef = useRef<string | null>(null)
  useEffect(() => {
    if (!pendingOfficeImport || officeImportRef.current === pendingOfficeImport) return
    officeImportRef.current = pendingOfficeImport
    const url = pendingOfficeImport
    setPendingOfficeImport(null)
    setOfficeSwitchBusy(office.circle || office.zone || null)
    void importFromGoogleLink(url)
      .catch(() => {
        setTables([])
        setLastGoogleLink(null)
        setOfficeImportPrompt(office.circle || office.zone || null)
      })
      .finally(() => setOfficeSwitchBusy(null))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pendingOfficeImport, office])

  return (
    <>
      <Skyline />
      <UpdateBanner />
      <ProfileMenu
        onLogout={() => {
          // Logging out unmounts this component, which would otherwise
          // clearTimeout() the debounced save below and drop whatever was
          // just typed (e.g. MB Scrutiny data entered seconds earlier).
          // Flush it synchronously first so it's written before that happens.
          pendingSaveRef.current?.()
          onLogout()
        }}
      />
      <div className="shell">
      <Sidebar
        active={tab}
        onSelect={setTab}
        onReset={resetToTab}
        tableCount={tables.length}
        unresolved={unresolved.length}
        createdDocCount={createdDocuments.length}
        office={office}
        onOfficeChange={changeOffice}
        onShowWhatsNew={() => setWhatsNew(CHANGELOG)}
      />

      <main className="workspace">
        <Fragment key={resetNonce}>
        <KeepAlive active={tab === 'dashboard'} mounted={mountedTabs.has('dashboard')}>
          <section className="page wide">
            <div className="page-head">
              <div className="page-ic violet">
                <IconCalendar />
              </div>
              <div className="page-head-text">
                <h1>Calendar</h1>
                <p>Telangana Government holiday calendar{calendar ? ` for ${calendar.year}` : ''}.</p>
              </div>
              <div className="page-head-action">
                <TenderNoticeButton
                  tables={tables}
                  office={office}
                  onGenerated={addTenderReminder}
                  onBidBatch={addBidBatch}
                />
              </div>
            </div>
            <Dashboard cached={calendar} onData={setCalendar} office={office} />
            <SeBidDocumentTile tables={tables} office={office} />
            <BidDocumentsPanel batches={officeBidBatches} onRemove={removeBidBatch} onUpdateWork={updateBidWork} />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'data'} mounted={mountedTabs.has('data')}>
          <section className="page wide">
            <div className="page-head">
              <div className="page-ic green">
                <IconTable />
              </div>
              <div className="page-head-text">
                <h1>Works List</h1>
              </div>
              <div className="page-head-action">
                {currentTable && (
                  <WorksListL1Update
                    // Remount on office change so its own "Updated X, Y
                    // unmatched" result banner (internal state, has no office
                    // awareness of its own) doesn't keep showing a stale
                    // result from the office just switched away from.
                    key={currentOfficeKey}
                    table={currentTable}
                    onChange={updateTable}
                    onUpdated={(rows, message) => setWorksFlash({ token: Date.now(), rows, message })}
                  />
                )}
                {currentTable && (
                  <button className="ghost" onClick={exportWorksList} title="Download as Excel">
                    <IconDownload /> Download
                  </button>
                )}
                {lastGoogleLink && (
                  <button className="ghost" onClick={refreshWorksList} disabled={refreshingWorks}>
                    <IconRefresh /> {refreshingWorks ? 'Refreshing…' : 'Refresh'}
                  </button>
                )}
                {worksListErrors.length > 0 && (
                  <button
                    type="button"
                    className="works-errors-btn blink"
                    title={`${worksListErrors.length} issue${worksListErrors.length === 1 ? '' : 's'} found in the Works List`}
                    onClick={() => setShowWincodeViolations(true)}
                  >
                    <IconWarn /> Errors ({worksListErrors.length})
                  </button>
                )}
              </div>
            </div>
            <QcPartiesEditor
              parties={currentOfficeKey ? qcParties[currentOfficeKey] : undefined}
              circleSelected={!!office.circle}
              onChange={(next) =>
                currentOfficeKey && setQcParties((prev) => ({ ...prev, [currentOfficeKey]: next }))
              }
            />
            {officeSwitchBusy && (
              <div className="notice warn office-switch-busy">
                <IconRefresh className="spin" /> Loading <strong>{officeSwitchBusy}</strong>'s Works List…
              </div>
            )}
            {officeImportPrompt && (
              <div className="notice warn">
                <IconWarn /> You changed your office to <strong>{officeImportPrompt}</strong>. Paste that office's Works
                List link below and Import to load its database.
              </div>
            )}
            {refreshError && <div className="notice error">{refreshError}</div>}
            {refreshSummary && <div className="notice ok">{refreshSummary}</div>}
            <GoogleLinkImport onImport={importFromGoogleLink} />
            {circleMismatchConfirm &&
              createPortal(
                <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => resolveCircleMismatch(false))}>
                  <div className="confirm-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
                    <div className="confirm-ic">
                      <IconWarn />
                    </div>
                    <h3>Works from another circle</h3>
                    <p className="confirm-warn">
                      <strong>{circleMismatchConfirm.count}</strong> of {circleMismatchConfirm.total} works in this list
                      don't belong to your office's Zone/Circle (
                      {[circleMismatchConfirm.zone, circleMismatchConfirm.circle].filter(Boolean).join(' / ')}).
                    </p>
                    <p className="confirm-hint">
                      Are you sure these works belong to your circle? (A circle reorganisation can leave some works still
                      tagged with their old circle — if that's the case here, Continue is safe.)
                    </p>
                    <div className="confirm-actions">
                      <button className="ghost" onClick={() => resolveCircleMismatch(false)}>
                        Don't Import
                      </button>
                      <button className="primary" onClick={() => resolveCircleMismatch(true)}>
                        Continue
                      </button>
                    </div>
                  </div>
                </div>,
                document.body
              )}
            {showWincodeViolations &&
              createPortal(
                <div className="editor-overlay" onMouseDown={closeOnBackdropMouseDown(() => setShowWincodeViolations(false))}>
                  <div className="confirm-modal" role="dialog" aria-modal="true" onClick={(e) => e.stopPropagation()}>
                    <div className="confirm-ic">
                      <IconWarn />
                    </div>
                    <h3>
                      {worksListErrors.length} issue{worksListErrors.length === 1 ? '' : 's'} found in the Works List
                    </h3>
                    <p className="confirm-hint">Data problems found while scanning the Works List — see each row below.</p>
                    <ul className="works-list-violation-list">
                      {worksListErrors.map((v) => (
                        <li key={`${v.type}:${v.key}`}>
                          {v.message} (row{v.rowIndices.length > 1 ? 's' : ''} {v.rowIndices.map((i) => i + 1).join(', ')})
                        </li>
                      ))}
                    </ul>
                    <div className="confirm-actions">
                      <button className="primary" onClick={() => setShowWincodeViolations(false)}>
                        Close
                      </button>
                    </div>
                  </div>
                </div>,
                document.body
              )}
            {currentTable ? (
              <>
                <ExcelInline
                  key={`${currentTable.id}:${worksFlash?.token ?? 0}`}
                  table={currentTable}
                  onChange={updateTable}
                  autofillRow={autofillWorksRowForLogin}
                  flashRows={worksFlash?.rows}
                  flashMessage={worksFlash?.message}
                />
              </>
            ) : (
              <div className="card">
                <div className="empty">
                  <IconTable />
                  <p>No works database yet. Create one with the standard work columns and start adding rows.</p>
                  <button className="primary" onClick={createDatabase}>
                    <IconPlus /> Create works database
                  </button>
                </div>
              </div>
            )}
            <CollisionPanel
              collisions={collisions}
              resolution={resolution}
              onResolve={resolveCollision}
            />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'printDoc'} mounted={mountedTabs.has('printDoc')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic sky">
                <IconPrint />
              </div>
              <div className="page-head-text">
                <h1>Issue Documents</h1>
              </div>
            </div>
            <PrintDocumentTab
              tables={tables}
              documents={bakedDocuments}
              onChange={reorderDocuments}
              onGoToWorksList={() => setTab('data')}
              office={office}
              qcParties={currentOfficeKey ? qcParties[currentOfficeKey] : undefined}
              theme={theme}
            />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'estimateWorkspace'} mounted={mountedTabs.has('estimateWorkspace')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic amber">
                <IconDoc />
              </div>
              <div className="page-head-text">
                <h1>BOQ, Schedule A, Deviation & Material Quantity</h1>
                <p>Upload an estimate once to download its BOQ, Schedule A, Deviation Statement, and Material Quantity.</p>
              </div>
            </div>
            <EstimateWorkspaceTab tables={tables} onChange={updateTable} office={office} />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'techSanction'} mounted={mountedTabs.has('techSanction')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconCheck />
              </div>
              <div className="page-head-text">
                <h1>Give Technical Sanction (Experimental)</h1>
                <p>Upload the Data Sheet and an Estimate to auto-fill rates from the Data Sheet.</p>
              </div>
            </div>
            <div className="notice warn">
              <IconWarn /> This is still an experimental feature and may not work as intended.
            </div>
            <GiveTechnicalSanctionTab />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'intimation'} mounted={mountedTabs.has('intimation')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic rose">
                <IconBell />
              </div>
              <div className="page-head-text">
                <h1>Intimation</h1>
                <p>
                  {intimationSubTab === 'intimation'
                    ? "Upload a PDF (or photos) plus an existing Intimation format to auto-fill it from what's in the document."
                    : 'Upload the portal’s “View Bidders” PDF to issue the Bid Capacity Evaluation Sheet — one column per participating bidder.'}
                </p>
              </div>
              {intimationSubTab === 'intimation' && (
                <div className="page-head-action">
                  <button
                    className="ghost"
                    onClick={() => setIntimationInstanceKey((k) => k + 1)}
                    title="Clear the uploaded documents and filled fields to start a new work"
                  >
                    <IconRefresh /> Clear
                  </button>
                  <div className="header-action-slot" ref={intimationHeaderActionRef} />
                </div>
              )}
            </div>
            <div className="doc-tabs">
              <button
                className={`doc-tab ${intimationSubTab === 'intimation' ? 'active' : ''}`}
                onClick={() => setIntimationSubTab('intimation')}
              >
                Issue Intimation
              </button>
              <button
                className={`doc-tab ${intimationSubTab === 'evaluation' ? 'active' : ''}`}
                onClick={() => setIntimationSubTab('evaluation')}
              >
                Evaluation Sheet
              </button>
            </div>
            {intimationSubTab === 'intimation' ? (
              <GiveIntimationTab
                key={intimationInstanceKey}
                tables={tables}
                onChange={updateTable}
                office={office}
                headerActionRef={intimationHeaderActionRef}
                theme={theme}
              />
            ) : (
              <EvaluationSheetTab office={office} />
            )}
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'workOrder'} mounted={mountedTabs.has('workOrder')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconClipboard />
              </div>
              <div className="page-head-text">
                <h1>Agreement and Work order</h1>
                <p>Fill the Work Order and Agreement for a work, and generate its Schedule A from the estimate / BOQ.</p>
              </div>
              <div className="page-head-action">
                <button
                  className="ghost"
                  onClick={() => setWorkOrderInstanceKey((k) => k + 1)}
                  title="Clear the uploaded documents and filled fields to start a new work"
                >
                  <IconRefresh /> Clear
                </button>
                <div className="header-action-slot" ref={workOrderHeaderActionRef} />
              </div>
            </div>
            <WorkOrderAgreementTab
              key={workOrderInstanceKey}
              tables={tables}
              onChange={updateTable}
              zoneLogin={!!loginZone && !loginCircle}
              office={office}
              headerActionRef={workOrderHeaderActionRef}
              theme={theme}
            />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'issueNotices'} mounted={mountedTabs.has('issueNotices')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconWarn />
              </div>
              <div className="page-head-text">
                <h1>Issue Notices</h1>
                <p>Upload the Online Intimation and L1 sheet to fill the Notice for an L-1 bidder who hasn't concluded the agreement.</p>
              </div>
              <div className="page-head-action">
                <button
                  className="ghost"
                  onClick={() => setIssueNoticesInstanceKey((k) => k + 1)}
                  title="Clear the uploaded documents and filled fields to start a new work"
                >
                  <IconRefresh /> Clear
                </button>
              </div>
            </div>
            <IssueNoticesTab key={issueNoticesInstanceKey} office={office} theme={theme} />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'mbScrutiny'} mounted={mountedTabs.has('mbScrutiny')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconEye />
              </div>
              <div className="page-head-text">
                <h1>MB Scrutiny list{officeLabel ? ` — ${officeLabel}` : ''}</h1>
                <p>Log each Measurement Book received for scrutiny and track it through to completion. This register is specific to the selected office.</p>
              </div>
              <div className="page-head-action">
                {officeMb.length > 0 && (
                  <button className="ghost" onClick={exportMbScrutiny} title="Download as Excel">
                    <IconDownload /> Download
                  </button>
                )}
                <button
                  className="ghost"
                  onClick={() => void restoreMbScrutiny()}
                  title="Restore from a previously downloaded Excel file"
                >
                  <IconUpload /> Restore
                </button>
              </div>
            </div>
            <MbScrutinyList items={officeMb} onChange={setOfficeMb} />

            <div className="settings-template-section" style={{ marginTop: 32 }}>
              <div className="settings-section-head">
                <button
                  type="button"
                  className="settings-section-toggle"
                  onClick={() => setMbMeasurementOpen((o) => !o)}
                  aria-expanded={mbMeasurementOpen}
                >
                  <IconChevronRight className={`settings-section-chevron ${mbMeasurementOpen ? 'open' : ''}`} />
                  <span className="settings-section-toggle-titles">
                    <h3 className="settings-template-section-title">Measurement Sheet Extraction</h3>
                    <p className="sub">
                      Upload photos or a scanned PDF of an MB measurement sheet to read it into an editable,
                      downloadable Excel grid.
                    </p>
                  </span>
                </button>
              </div>
              <div style={{ display: mbMeasurementOpen ? undefined : 'none' }}>
                <MbMeasurementUploadTab />
              </div>
            </div>
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'search'} mounted={mountedTabs.has('search')}>
          <section className="page wide">
            <div className="page-head">
              <div className="page-ic green">
                <IconSearch />
              </div>
              <div className="page-head-text">
                <h1>Search Tender</h1>
                <p>Search live tenders from the Telangana e-procurement portal by ID, NIT no, or work name.</p>
              </div>
            </div>
            <SearchTender onAddReminder={addTenderReminderFromRow} />
            <TenderReminders
              reminders={tenderReminders}
              onDelete={removeTenderReminder}
              onRefresh={refreshTenderReminder}
              refreshingId={refreshingReminderId}
            />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'todo'} mounted={mountedTabs.has('todo')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconChecklist />
              </div>
              <div className="page-head-text">
                <h1>To Do List{officeLabel ? ` — ${officeLabel}` : ''}</h1>
                <p>
                  Track tasks with a target completion date. Ticking one off keeps it visible for
                  the rest of the day, then it rolls off the list. This list is specific to the
                  selected office.
                </p>
              </div>
            </div>
            <TodoList todos={officeTodos} onChange={setOfficeTodos} />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'tools'} mounted={mountedTabs.has('tools')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic teal">
                <IconTools />
              </div>
              <div className="page-head-text">
                <h1>Tools</h1>
              </div>
            </div>
            <ToolsTab tables={tables} onChange={updateTable} office={office} documents={bakedDocuments} theme={theme} />
          </section>
        </KeepAlive>

        <KeepAlive active={tab === 'settings'} mounted={mountedTabs.has('settings')}>
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconSettings />
              </div>
              <div className="page-head-text">
                <h1>Settings</h1>
              </div>
            </div>
            <SettingsTab office={office} theme={theme} onThemeChange={changeTheme} />
          </section>
        </KeepAlive>
        </Fragment>
      </main>
      </div>
      {whatsNew.length > 0 && (
        <WhatsNew
          entries={whatsNew}
          onClose={() => {
            // Record the running version as seen only now, so closing is what
            // dismisses it for good (it re-shows until the user acknowledges).
            void api.getAppVersion().then((v) => v && localStorage.setItem(LAST_SEEN_VERSION_KEY, v))
            setWhatsNew([])
          }}
        />
      )}
    </>
  )
}

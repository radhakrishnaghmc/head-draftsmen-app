import { useEffect, useMemo, useRef, useState } from 'react'
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
import { entriesOf, corporationByName } from './zoneCircleDirectory'
import { type Office, officeKey } from './office'
import { matchPlaceholdersToColumns } from '@core/createDocument'
import { mergeTables } from '@core/merge'
import Sidebar, { type TabKey } from './components/Sidebar'
import OfficeSelector from './components/OfficeSelector'
import ExcelInline from './components/ExcelInline'
import CollisionPanel from './components/CollisionPanel'
import SearchTender from './components/SearchTender'
import Dashboard from './components/Dashboard'
import Skyline from './components/Skyline'
import ProfileMenu from './components/ProfileMenu'
import UpdateBanner from './components/UpdateBanner'
import TenderNoticeButton from './components/TenderNoticeButton'
import BidDocumentsPanel from './components/BidDocumentsPanel'
import GoogleLinkImport from './components/GoogleLinkImport'
import WorksListL1Update from './components/WorksListL1Update'
import EstimateWorkspaceTab from './components/EstimateWorkspaceTab'
import GiveTechnicalSanctionTab from './components/GiveTechnicalSanctionTab'
import GiveIntimationTab from './components/GiveIntimationTab'
import WorkOrderAgreementTab from './components/WorkOrderAgreementTab'
import PrintDocumentTab from './components/PrintDocumentTab'
import ToolsTab from './components/ToolsTab'
import TodoList from './components/TodoList'
import MbScrutinyList from './components/MbScrutinyList'
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
  IconTools,
  IconWarn
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
  BidDocumentBatch
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

interface Props {
  onLogout: () => void
  /** The Head Draughtsman's chosen office (Corporation/Zone/Circle), selected in the sidebar — drives document prep and Works List validation. */
  office: Office
  onOfficeChange: (office: Office) => void
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

  const [tab, setTab] = useState<TabKey>('dashboard')
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
  // Works List link remembered per office (key: officeKey) — synced, so a
  // circle's database reloads on return and follows the user across systems.
  const [worksListLinks, setWorksListLinks] = useState<Record<string, string>>({})
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

  // Set for exactly the next autosave whenever state was just applied FROM a
  // remote sync — so that save persists locally but is NOT pushed back to the
  // cloud. Without this, receiving a remote change re-pushes it, the other
  // sessions re-push it in turn, and the resulting write storm lets a stale
  // copy overwrite (e.g.) a To Do task another session just added.
  const savedFromRemoteRef = useRef(false)

  // The current Excel shown on the Data tab (single-workbook workflow).
  const currentTable = tables[0] ?? null

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
            if (mismatches.length > 0) {
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
    const handle = setTimeout(() => {
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
          resolution,
          todos,
          lastGoogleLink: lastGoogleLink ?? undefined,
          worksListLinks,
          tenderReminders,
          createdDocuments,
          seededDocVersion,
          bidDocumentBatches,
          mbScrutiny
        },
        skipCloud
      )
    }, 400)
    return () => clearTimeout(handle)
  }, [
    hydrated,
    blockedWorksList,
    tables,
    tablesByOffice,
    currentOfficeKey,
    resolution,
    todos,
    lastGoogleLink,
    worksListLinks,
    tenderReminders,
    createdDocuments,
    seededDocVersion,
    bidDocumentBatches,
    mbScrutiny
  ])

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
    if (mbScrutiny.length === 0) return
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
    const rows = [...mbScrutiny]
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

    // Only a works list belonging to the logged-in Head Draughtsman's own
    // Zone/Circle is accepted — a row explicitly tagged with a different
    // Zone/Circle, or whose "Name of the work" names a different one,
    // rejects the whole import rather than silently mixing works lists.
    let result: ExcelTable
    if (loginZone && loginCircle) {
      const { table: checked, mismatches } = enforceZoneCircle(normalized, loginZone, loginCircle, officeEntries)
      if (mismatches.length > 0) {
        const examples = mismatches
          .slice(0, 3)
          .map((m) => `"${m.workName || `Row ${m.rowIndex + 1}`}" (${[m.foundZone, m.foundCircle].filter(Boolean).join(' / ')})`)
          .join(', ')
        throw new Error(
          `This works list doesn't match your office's Zone/Circle (${loginZone} / ${loginCircle}). ` +
            `${mismatches.length} row${mismatches.length === 1 ? '' : 's'} conflict: ${examples}` +
            `${mismatches.length > 3 ? ', …' : ''}.`
        )
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
    setBidDocumentBatches((prev) => [...prev, batch])
  }

  function removeBidBatch(id: string) {
    setBidDocumentBatches((prev) => prev.filter((b) => b.id !== id))
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
    void importFromGoogleLink(url).catch(() => {
      setTables([])
      setLastGoogleLink(null)
      setOfficeImportPrompt(office.circle || office.zone || null)
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pendingOfficeImport, office])

  return (
    <>
      <Skyline />
      <UpdateBanner />
      <ProfileMenu onLogout={onLogout} />
      <div className="shell">
      <Sidebar
        active={tab}
        onSelect={setTab}
        tableCount={tables.length}
        unresolved={unresolved.length}
        createdDocCount={createdDocuments.length}
        office={office}
        onOfficeChange={changeOffice}
      />

      <main className="workspace">
        {tab === 'dashboard' && (
          <section className="page wide">
            <div className="page-head">
              <div className="page-ic violet">
                <IconCalendar />
              </div>
              <div className="page-head-text">
                <h1>Calendar</h1>
                <p>Telangana Government holiday calendar for 2026.</p>
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
            <Dashboard cached={calendar} onData={setCalendar} />
            <BidDocumentsPanel batches={bidDocumentBatches} onRemove={removeBidBatch} />
          </section>
        )}

        {tab === 'data' && (
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
              </div>
            </div>
            <OfficeSelector office={office} onChange={changeOffice} />
            {officeImportPrompt && (
              <div className="notice warn">
                <IconWarn /> You changed your office to <strong>{officeImportPrompt}</strong>. Paste that office's Works
                List link below and Import to load its database.
              </div>
            )}
            {refreshError && <div className="notice error">{refreshError}</div>}
            {refreshSummary && <div className="notice ok">{refreshSummary}</div>}
            <GoogleLinkImport onImport={importFromGoogleLink} />
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
        )}

        {tab === 'printDoc' && (
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
            />
          </section>
        )}

        {tab === 'estimateWorkspace' && (
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
        )}

        {tab === 'techSanction' && (
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
        )}

        {tab === 'intimation' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic rose">
                <IconBell />
              </div>
              <div className="page-head-text">
                <h1>Intimation</h1>
                <p>Upload a PDF (or photos) plus an existing Intimation format to auto-fill it from what's in the document.</p>
              </div>
            </div>
            <GiveIntimationTab tables={tables} onChange={updateTable} office={office} />
          </section>
        )}

        {tab === 'workOrder' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconClipboard />
              </div>
              <div className="page-head-text">
                <h1>Agreement and Work order</h1>
                <p>Fill the Work Order and Agreement for a work, and generate its Schedule A from the estimate / BOQ.</p>
              </div>
            </div>
            <WorkOrderAgreementTab
              tables={tables}
              onChange={updateTable}
              zoneLogin={!!loginZone && !loginCircle}
              office={office}
            />
          </section>
        )}

        {tab === 'mbScrutiny' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconEye />
              </div>
              <div className="page-head-text">
                <h1>MB Scrutiny list</h1>
                <p>Log each Measurement Book received for scrutiny and track it through to completion.</p>
              </div>
              {mbScrutiny.length > 0 && (
                <div className="page-head-action">
                  <button className="ghost" onClick={exportMbScrutiny} title="Download as Excel">
                    <IconDownload /> Download
                  </button>
                </div>
              )}
            </div>
            <MbScrutinyList items={mbScrutiny} onChange={setMbScrutiny} />
          </section>
        )}

        {tab === 'search' && (
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
        )}

        {tab === 'todo' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconChecklist />
              </div>
              <div className="page-head-text">
                <h1>To Do List</h1>
                <p>
                  Track tasks with a target completion date. Ticking one off keeps it visible for
                  the rest of the day, then it rolls off the list.
                </p>
              </div>
            </div>
            <TodoList todos={todos} onChange={setTodos} />
          </section>
        )}

        {tab === 'tools' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic teal">
                <IconTools />
              </div>
              <div className="page-head-text">
                <h1>Tools</h1>
              </div>
            </div>
            <ToolsTab tables={tables} onChange={updateTable} office={office} documents={bakedDocuments} />
          </section>
        )}
      </main>
      </div>
    </>
  )
}

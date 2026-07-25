import { useEffect, useMemo, useRef, useState } from 'react'
import { api } from './ipc'
import { prefetchTenders, fetchTenders } from './tenderCache'
import { createWorksTable, applyWorksSchema, applyWorksSchemaWithMapping, WORKS_COLUMNS } from './worksSchema'
import { syncWorksListFromTenderPortal } from './worksListTenderSync'
import { enforceZoneCircle, fillCircleNumber } from './zoneCircleCheck'
import { matchPlaceholdersToColumns } from '@core/createDocument'
import { findCircleSheet, mergeMonitoringRows } from '@core/monitoringImport'
import { guessHeaderRow, buildTableFromGrid } from '@core/sheet'
import { mergeTables } from '@core/merge'
import Sidebar, { type TabKey } from './components/Sidebar'
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
import MonitoringLinkImport, { type MonitoringMergeSummary } from './components/MonitoringLinkImport'
import EstimateToBoqTab from './components/EstimateToBoqTab'
import ScheduleATab from './components/ScheduleATab'
import GetDeviationTab from './components/GetDeviationTab'
import GiveTechnicalSanctionTab from './components/GiveTechnicalSanctionTab'
import CreateDocumentTab from './components/CreateDocumentTab'
import PrintDocumentTab from './components/PrintDocumentTab'
import UploadPhotosTab from './components/UploadPhotosTab'
import MaterialQuantityTab from './components/MaterialQuantityTab'
import TodoList from './components/TodoList'
import TenderReminders from './components/TenderReminders'
import {
  IconTable,
  IconDoc,
  IconPrint,
  IconSearch,
  IconCalendar,
  IconPlus,
  IconChecklist,
  IconRefresh,
  IconCheck,
  IconImage
} from './components/Icons'
import type {
  ExcelTable,
  MergedDataset,
  CollisionResolution,
  TodoItem,
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

interface Props {
  onLogout: () => void
  /** The logged-in Head Draftsman's own Zone/Circle, from the login credentials sheet — shown in the sidebar. */
  loginZone?: string
  loginCircle?: string
  /** Only present for a Circle-level login (not a Zone-level one, which spans many circles and has no single number). */
  loginCircleNumber?: string
}

export default function App({ onLogout, loginZone, loginCircle, loginCircleNumber }: Props) {
  const [tab, setTab] = useState<TabKey>('dashboard')
  const [calendar, setCalendar] = useState<CalendarData | null>(null)

  const [tables, setTables] = useState<ExcelTable[]>([])
  const [resolution, setResolution] = useState<CollisionResolution>({})

  const [todos, setTodos] = useState<TodoItem[]>([])
  const [lastGoogleLink, setLastGoogleLink] = useState<string | null>(null)
  const [refreshingWorks, setRefreshingWorks] = useState(false)
  const [refreshError, setRefreshError] = useState<string | null>(null)
  // Progress/outcome of the tender-portal sync that follows the Google-link
  // refresh — see refreshWorksList and src/worksListTenderSync.ts.
  const [tenderSyncStatus, setTenderSyncStatus] = useState<string | null>(null)
  const [tenderSyncSummary, setTenderSyncSummary] = useState<string | null>(null)
  const [tenderReminders, setTenderReminders] = useState<TenderReminder[]>([])
  const [refreshingReminderId, setRefreshingReminderId] = useState<string | null>(null)
  const [createdDocuments, setCreatedDocuments] = useState<CreatedDocument[]>([])
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
  // Set when a document is sent from Issue Document to Create New Document
  // for edits — cleared once loaded/saved/cancelled there.
  const [editingDoc, setEditingDoc] = useState<CreatedDocument | null>(null)

  function editCreatedDocument(doc: CreatedDocument) {
    setEditingDoc(doc)
    setTab('createDoc')
  }

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
    return Promise.all(
      docs.map(async (d) => ({ ...d, docx: await api.bakeFixedPlaceholdersInDocument(d.docx, values) }))
    )
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
          const loadedTables = (s.tables ?? []).map(applyWorksSchema)

          // A previously-saved Works List belonging to a different Zone/Circle
          // than the one logged in now must not be shown — same rule as a
          // fresh import (see importFromGoogleLink), applied to what's already
          // on disk too.
          const first = loadedTables[0]
          if (loginZone && loginCircle && first) {
            const { table: checked, mismatches } = enforceZoneCircle(first, loginZone, loginCircle)
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
          setLastGoogleLink(s.lastGoogleLink ?? null)
          setTenderReminders((s.tenderReminders ?? []).map(migrateTenderReminder))
          setCreatedDocuments(await bakeLoginPlaceholders(s.createdDocuments ?? []))
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
      if (partial.tables) {
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
      if (partial.lastGoogleLink !== undefined) setLastGoogleLink(partial.lastGoogleLink ?? null)
      if (partial.tenderReminders) setTenderReminders(partial.tenderReminders.map(migrateTenderReminder))
      if (partial.createdDocuments) setCreatedDocuments(await bakeLoginPlaceholders(partial.createdDocuments))
      if (partial.bidDocumentBatches) setBidDocumentBatches(partial.bidDocumentBatches)
    })
  }, [hydrated, blockedWorksList])

  // ── Persistence: save the workspace whenever it changes (debounced) ─
  // While a Works List is withheld for a Zone/Circle mismatch
  // (blockedWorksList), `tables` is deliberately empty in-memory — so we
  // substitute the real on-disk data (withheldTablesRef) back in for that
  // one field. Everything else (documents, todos, reminders, ...) always
  // saves from live state, so it's never paused or lost.
  useEffect(() => {
    if (!hydrated) return
    const handle = setTimeout(() => {
      api.saveState({
        version: 1,
        tables: blockedWorksList ? (withheldTablesRef.current ?? tables) : tables,
        resolution,
        todos,
        lastGoogleLink: lastGoogleLink ?? undefined,
        tenderReminders,
        createdDocuments,
        bidDocumentBatches
      })
    }, 400)
    return () => clearTimeout(handle)
  }, [
    hydrated,
    blockedWorksList,
    tables,
    resolution,
    todos,
    lastGoogleLink,
    tenderReminders,
    createdDocuments,
    bidDocumentBatches
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
    const mapping = matchPlaceholdersToColumns(WORKS_COLUMNS, imported.headers, embeddings)

    const normalized = applyWorksSchemaWithMapping(imported.headers, rows, mapping, {
      id: `works-${Date.now()}`,
      name: 'Works database',
      path: ''
    })

    // Only a works list belonging to the logged-in Head Draftsman's own
    // Zone/Circle is accepted — a row explicitly tagged with a different
    // Zone/Circle, or whose "Name of the work" names a different one,
    // rejects the whole import rather than silently mixing works lists.
    let result: ExcelTable
    if (loginZone && loginCircle) {
      const { table: checked, mismatches } = enforceZoneCircle(normalized, loginZone, loginCircle)
      if (mismatches.length > 0) {
        const examples = mismatches
          .slice(0, 3)
          .map((m) => `"${m.workName || `Row ${m.rowIndex + 1}`}" (${[m.foundZone, m.foundCircle].filter(Boolean).join(' / ')})`)
          .join(', ')
        throw new Error(
          `This works list doesn't match your login's Zone/Circle (${loginZone} / ${loginCircle}). ` +
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

    return { added: rows.length, table: result }
  }

  // A monitoring format is one workbook with a sheet per circle — pick the
  // one matching this login's own Circle and merge it into the existing
  // Works List (fill blanks on matched rows, add rows for new works), rather
  // than replacing the database the way the plain link import above does.
  async function importFromMonitoringLink(url: string): Promise<MonitoringMergeSummary> {
    if (!loginCircle) {
      throw new Error(
        "Your login isn't tied to a single Circle, so there's no one circle sheet to pick from a monitoring format link."
      )
    }
    const table = tables[0]
    if (!table) {
      throw new Error("Add a Works List first — there's nothing yet to merge the monitoring data into.")
    }

    const sheets = await api.importAllSheetsFromLink(url)
    const picked = findCircleSheet(sheets, loginCircle, loginCircleNumber)
    if (!picked) {
      throw new Error(`Couldn't find a sheet matching your Circle ("${loginCircle}") in that monitoring format link.`)
    }
    const monitoringTable = buildTableFromGrid(picked.grid, guessHeaderRow(picked.grid), {
      id: 'monitoring',
      name: picked.sheetName,
      path: ''
    })
    if (monitoringTable.rows.length === 0) {
      throw new Error(`The "${picked.sheetName}" sheet has no usable rows.`)
    }

    let embeddings: { labelVectors: number[][]; columnVectors: number[][] } | undefined
    try {
      const [labelVectors, columnVectors] = await Promise.all([
        api.embedTexts(table.headers),
        api.embedTexts(monitoringTable.headers)
      ])
      embeddings = { labelVectors, columnVectors }
    } catch {
      // Neural matching unavailable — matchPlaceholdersToColumns falls back
      // to plain token overlap automatically when no embeddings are passed.
      embeddings = undefined
    }
    const mapping = matchPlaceholdersToColumns(table.headers, monitoringTable.headers, embeddings)
    const result = mergeMonitoringRows(table, monitoringTable.rows, mapping)

    setTables([fillCircleNumber(result.table, loginCircleNumber), ...tables.slice(1)])
    return {
      matched: result.matchedCount,
      added: result.addedCount,
      filled: result.filledCount,
      sheetName: picked.sheetName
    }
  }

  // Re-pull the Works List from whichever Google link it was last filled
  // from, so the user doesn't have to re-paste the URL to get fresh data —
  // then, against that freshly-pulled table, search the tender portal by
  // each of its Circles and fill ECV/Tender Notice No/Tender ID from
  // whichever tender matches each work by name (src/worksListTenderSync.ts).
  // Uses the table importFromGoogleLink itself just produced, not the
  // `tables` state variable — that state update hasn't necessarily been
  // committed yet by the time this line runs.
  async function refreshWorksList() {
    if (!lastGoogleLink) return
    setRefreshingWorks(true)
    setRefreshError(null)
    setTenderSyncSummary(null)
    try {
      const { table } = await importFromGoogleLink(lastGoogleLink)
      try {
        const { table: synced, matchedCount, tenderCount, circleCount } = await syncWorksListFromTenderPortal(
          table,
          (p) => setTenderSyncStatus(`Searching tenders for ${p.circle} (${p.circleIndex + 1} of ${p.circleCount})…`)
        )
        setTables([synced])
        setTenderSyncSummary(
          circleCount === 0
            ? 'Refreshed from the Google link — no Circle values on the Works List to search tenders for.'
            : `Refreshed from the Google link, then matched ${matchedCount} work${matchedCount === 1 ? '' : 's'} against ${tenderCount} tender${tenderCount === 1 ? '' : 's'} found across ${circleCount} circle${circleCount === 1 ? '' : 's'}.`
        )
      } catch (e) {
        // The Google-link refresh itself already succeeded — a failed tender
        // sync on top of it is a partial-success notice, not a full error.
        setTenderSyncSummary(
          `Refreshed from the Google link, but the tender-portal search failed: ${e instanceof Error ? e.message : String(e)}`
        )
      }
    } catch (e) {
      setRefreshError(e instanceof Error ? e.message : String(e))
    } finally {
      setRefreshingWorks(false)
      setTenderSyncStatus(null)
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
        zone={loginZone}
        circle={loginCircle}
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
                <TenderNoticeButton tables={tables} onGenerated={addTenderReminder} onBidBatch={addBidBatch} />
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
              {lastGoogleLink && (
                <div className="page-head-action">
                  <button className="ghost" onClick={refreshWorksList} disabled={refreshingWorks}>
                    <IconRefresh /> {refreshingWorks ? 'Refreshing…' : 'Refresh'}
                  </button>
                </div>
              )}
            </div>
            {refreshError && <div className="notice error">{refreshError}</div>}
            {tenderSyncStatus && <div className="notice">{tenderSyncStatus}</div>}
            {tenderSyncSummary && <div className="notice ok">{tenderSyncSummary}</div>}
            <GoogleLinkImport onImport={importFromGoogleLink} />
            {loginCircle && <MonitoringLinkImport onImport={importFromMonitoringLink} />}
            {currentTable ? (
              <>
                <ExcelInline key={currentTable.id} table={currentTable} onChange={updateTable} />
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

        {tab === 'createDoc' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic rose">
                <IconDoc />
              </div>
              <div className="page-head-text">
                <h1>Create New Document</h1>
                <p>
                  Paste a document and add {'{{Placeholder}}'} markers — find it on the Issue
                  Document tab when you're ready to fill and print it.
                </p>
              </div>
            </div>
            <CreateDocumentTab
              documents={createdDocuments}
              onChange={setCreatedDocuments}
              editingDoc={editingDoc}
              onDoneEditing={() => setEditingDoc(null)}
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
                <h1>Issue Document</h1>
                <p>
                  Pick a saved document and a Works List row, then create the filled output —
                  matched to your columns by the neural model, not exact name matching.
                </p>
              </div>
            </div>
            <PrintDocumentTab
              tables={tables}
              documents={createdDocuments}
              onChange={setCreatedDocuments}
              onEdit={editCreatedDocument}
              onGoToWorksList={() => setTab('data')}
            />
          </section>
        )}

        {tab === 'boq' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic amber">
                <IconDoc />
              </div>
              <div className="page-head-text">
                <h1>Download BOQ</h1>
                <p>Upload an estimate to generate its Bill of Quantities.</p>
              </div>
            </div>
            <EstimateToBoqTab tables={tables} onChange={updateTable} />
          </section>
        )}

        {tab === 'scheduleA' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic teal">
                <IconDoc />
              </div>
              <div className="page-head-text">
                <h1>Download Schedule A</h1>
                <p>Upload a BOQ to generate its Schedule A.</p>
              </div>
            </div>
            <ScheduleATab tables={tables} />
          </section>
        )}

        {tab === 'deviation' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconDoc />
              </div>
              <div className="page-head-text">
                <h1>Get Deviation Statement</h1>
                <p>Upload an estimate to generate its Deviation Statement.</p>
              </div>
            </div>
            <GetDeviationTab tables={tables} />
          </section>
        )}

        {tab === 'techSanction' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconCheck />
              </div>
              <div className="page-head-text">
                <h1>Give Technical Sanction</h1>
                <p>Upload the Data Sheet and an Estimate to auto-fill rates from the Data Sheet.</p>
              </div>
            </div>
            <GiveTechnicalSanctionTab />
          </section>
        )}

        {tab === 'photoEstimate' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic rose">
                <IconImage />
              </div>
              <div className="page-head-text">
                <h1>Upload Photos to Get Estimate</h1>
                <p>Upload site photos and drag to arrange them in the order they should appear.</p>
              </div>
            </div>
            <UploadPhotosTab tables={tables} onChange={updateTable} />
          </section>
        )}

        {tab === 'materialQuantity' && (
          <section className="page">
            <div className="page-head">
              <div className="page-ic">
                <IconDoc />
              </div>
              <div className="page-head-text">
                <h1>Download Material Quantity</h1>
                <p>Upload an estimate to compute Stone Aggregates, Sand, Gravel, Granite/Napa Slabs, Cement, and Steel quantities.</p>
              </div>
            </div>
            <MaterialQuantityTab />
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
      </main>
      </div>
    </>
  )
}

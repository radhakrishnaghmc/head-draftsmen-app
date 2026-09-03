import { app, BrowserWindow, dialog, ipcMain, shell, screen } from 'electron'
import * as fs from 'fs'
import * as path from 'path'
import * as https from 'https'
import * as os from 'os'
import * as firebaseSync from './firebaseSync'
import { initAutoUpdate, restartToUpdate, checkForUpdatesManually } from './autoUpdate'
import { IPC } from './ipc-contract'
import type {
  ManualCheckResult,
  AgreementBundleFile,
  ActiveSessionInfo,
  PickedTenderDocument,
  VerifyDocItem,
  VerifyDocResult
} from './ipc-contract'
import { collectTenderDocuments } from './tenderDocumentScan'
import {
  workOrderTemplateFileName,
  agreementTemplateFileName,
  intimationTemplateFileName,
  fileBackerTemplateFileName,
  civilTenderTemplateFileName
} from '../core/workOrderTemplateVariants'
import { parseExcelFile, readExcelGrid, readAllSheetGrids, buildWorkbookBuffer, readSheetPreviews } from '../core/excel'
import type { SheetPreview } from '../core/excel'
import { recognizeImages } from './ocr'
import { extractMbMeasurementSheet } from './mbMeasurementOcr'
import { mbMeasurementRowsToGrid } from '../core/mbMeasurementExtract'
import { runSplitInWorker } from './splitRunner'
import { pdfPageCount, mergePdfFiles, splitPdfFile } from './pdfTools'
import { applyTechnicalSanctionEdits } from '../core/technicalSanctionOutput'
import type { CellEdit } from '../core/technicalSanction'
import { embedTexts, warmEmbeddings } from './embeddings'
import { buildScheduleAWorkbook, rowsToScheduleAItems } from '../core/scheduleA'
import type { ScheduleAMeta } from '../core/scheduleA'
import { fillScheduleATemplate, fillSeScheduleATemplate } from '../core/scheduleATemplate'
import { fillBoqTemplate, rowsToBoqData } from '../core/boqTemplate'
import { fillDeviationTemplate } from '../core/deviationTemplate'
import { buildEvaluationSheet } from '../core/evaluationSheet'
import type { EvaluationSheetInput } from '../core/evaluationSheet'
import type { DeviationItem, DeviationMeta } from '../core/deviationTemplate'
import { buildDetailedEstimateWorkbook } from '../core/estimateTemplate'
import type { DetailedEstimateMeta } from '../core/estimateTemplate'
import type { EstimateWorkItem } from '../core/estimateExtract'
import { fillMaterialTemplate } from '../core/materialTemplate'
import type { MaterialEstimateMeta } from '../core/materialTemplate'
import type { MaterialTotals } from '../core/materialEstimate'
import { parseCalendarHtml } from '../core/calendar'
import { importTableFromGoogleLink, importAllSheetsFromGoogleLink } from '../core/googleImport'
import { validateLogin } from '../core/auth'
import type { LoginResult } from '../core/auth'
import { rememberOfflineAuth, checkOfflineAuth } from './offlineAuth'
import { fillTenderNotice } from '../core/tenderNotice'
import type { TenderNoticeInput } from '../core/tenderNotice'
import { fillBidDocument, fillSeBidDocument } from '../core/bidDocument'
import type { BidDocumentInput } from '../core/bidDocument'
import type { CalendarData } from '../core/calendar'
import { convertHtmlToDocx } from '../core/htmlToDocx'
import { textToParagraphsHtml, buildPhotosWorkbook, buildWorkbookFromRows } from '../core/photosToDoc'
import { convertPdfToDocx } from '../core/pdfToDocx'
import { buildDocx, type DocBlock } from '../core/docxBuilder'
import type { OcrPage } from '../core/ocrReconstruct'
import { convertDocxToPdf, docxToPageImages } from '../core/docxToPdf'
import { mergeDocxBuffers } from '../core/mergeDocx'
import { splitDocxByPageBreaks } from '../core/splitDocx'
import { ocrGpsOverlay } from './gpsOcr'
import { fetchCementSteelRates, downloadCementSteelRateBuffer } from './cementSteelRates'
import type { CementSteelRate } from '../core/cementSteelRates'
import { sanitizeDocxForWord2007 } from '../core/word2007Compat'
import {
  listParagraphs,
  applyParagraphEdits,
  findPlaceholdersInAllParts,
  fillPlaceholdersInAllParts,
  bakeFixedPlaceholdersInDocx,
  extractAllText
} from '../core/docx-edit'
import {
  verifyPlaceholderCoverage,
  verifyAmountMath,
  verifyCorporationWording,
  verifyReservedTag,
  combineVerify
} from '../core/documentVerify'
import type { PlaceholderMatch } from '../core/createDocument'
import type {
  ExcelTable,
  TenderQuery,
  TenderResult,
  PersistedState
} from '../core/types'
import { pruneLegacyDocuments } from '../core/types'
import type { SheetGrid } from '../core/sheet'

let mainWindow: BrowserWindow | null = null

function createWindow(): void {
  // Clamp to the display's work area (screen minus the taskbar/dock), not
  // the raw screen size — on common 1366x768 Windows laptops a hardcoded
  // 1200x820 window is taller than the usable desktop, so its bottom edge
  // (dialog action buttons, e.g. Tender Notice's "Issue") renders behind
  // the taskbar and can't be clicked.
  const { width: waWidth, height: waHeight } = screen.getPrimaryDisplay().workAreaSize
  mainWindow = new BrowserWindow({
    width: Math.min(1200, waWidth),
    height: Math.min(820, waHeight),
    minWidth: 900,
    minHeight: Math.min(640, waHeight),
    title: 'Agreement Desk',
    backgroundColor: '#ffffff',
    // A packaged build already has the icon embedded into the native app
    // bundle/exe by electron-builder — this explicit path is only needed in
    // dev, where build/icon.png wouldn't otherwise be picked up.
    ...(!app.isPackaged ? { icon: path.join(__dirname, '../../build/icon.png') } : {}),
    webPreferences: {
      preload: path.join(__dirname, '../preload/index.js'),
      sandbox: false,
      contextIsolation: true,
      nodeIntegration: false,
      // Enables Chromium's built-in PDF viewer, used by the Photos → PDF tool's
      // preview (a blob: PDF shown in an <iframe>).
      plugins: true
    }
  })

  if (process.env['ELECTRON_RENDERER_URL']) {
    mainWindow.loadURL(process.env['ELECTRON_RENDERER_URL'])
  } else {
    mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'))
  }

  // Without this, `mainWindow` keeps pointing at a destroyed BrowserWindow
  // after the user closes it — any later async callback (e.g. a Firestore
  // sync update arriving after close) that calls mainWindow.webContents.send
  // would then throw "Object has been destroyed" as an uncaught exception,
  // crashing the whole app instead of just skipping the update.
  mainWindow.on('closed', () => {
    mainWindow = null
  })
}

app.whenReady().then(() => {
  registerHandlers()
  createWindow()
  initAutoUpdate(() => mainWindow)
  warmEmbeddings()
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

app.on('before-quit', () => {
  void firebaseSync.endSession()
})

function registerHandlers(): void {
  // The login whose workspace this machine is currently holding. The local
  // state cache (state.json) is namespaced by this, so on a SHARED computer a
  // different login never loads — and never leaks — the previous user's data.
  // Null only before any login (i.e. dev, where the login screen is skipped).
  let currentLoginId: string | null = null

  ipcMain.handle(IPC.pickExcels, async (): Promise<ExcelTable[]> => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: 'Select Excel data files',
      filters: [{ name: 'Excel workbooks', extensions: ['xlsx', 'xls'] }],
      properties: ['openFile', 'multiSelections']
    })
    if (result.canceled) return []
    return result.filePaths.map((p) => parseExcelFile(p))
  })

  ipcMain.handle(IPC.pickExcelGrids, async (): Promise<SheetGrid[]> => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: 'Select Excel data files',
      filters: [{ name: 'Excel workbooks', extensions: ['xlsx', 'xls'] }],
      properties: ['openFile', 'multiSelections']
    })
    if (result.canceled) return []
    // Read *every* sheet of each picked workbook (not just the first): a
    // workbook routinely holds one estimate per sheet, and the estimate flow
    // then splits any sheet that stacks several estimates. See
    // EstimateUploadTab's uploadEstimates and splitEstimateBlocks.
    return result.filePaths.flatMap((p) => readAllSheetGrids(p))
  })

  ipcMain.handle(IPC.pickEstimateGrid, async (): Promise<SheetGrid | null> => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: 'Select Estimate',
      filters: [{ name: 'Excel workbooks', extensions: ['xlsx', 'xls'] }],
      properties: ['openFile']
    })
    if (result.canceled || result.filePaths.length === 0) return null
    return readExcelGrid(result.filePaths[0])
  })

  ipcMain.handle(IPC.pickDataSheet, async (): Promise<SheetGrid[] | null> => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: 'Select Data Sheet (rates database)',
      filters: [{ name: 'Excel workbooks', extensions: ['xlsx', 'xls'] }],
      properties: ['openFile']
    })
    if (result.canceled || result.filePaths.length === 0) return null
    return readAllSheetGrids(result.filePaths[0])
  })

  ipcMain.handle(IPC.pickTenderDocuments, async (): Promise<PickedTenderDocument[]> => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: 'Select L1 sheets / Online Intimations — files, folders, or both',
      // macOS lets a user pick loose files and whole folders together in one
      // dialog; Windows' native picker can't mix the two modes and falls
      // back to whichever this resolves to there (still lets a Windows user
      // pick either files or a folder, just not both at once).
      properties: ['openFile', 'openDirectory', 'multiSelections']
    })
    if (result.canceled || result.filePaths.length === 0) return []
    return collectTenderDocuments(result.filePaths, (phase, done, total) => {
      mainWindow?.webContents.send(IPC.tenderScanProgress, { phase, done, total })
    })
  })

  ipcMain.handle(IPC.ocrEstimatePhotos, async (_e, dataUrls: string[]): Promise<SheetGrid> => {
    const buffers = dataUrls.map((u) => Buffer.from(u.split(',')[1] ?? '', 'base64'))
    // OCR all pages in parallel across the process pool (one per core), then
    // stitch them back together in page order.
    const perPage = await recognizeImages(buffers)
    const allRows: string[][] = []
    for (const lines of perPage) {
      // The OCR engine's detector already returns each printed line as one
      // clean unit of text (unlike a per-word engine, whose word boxes
      // would need reassembling into rows/columns by position) — sorting by
      // vertical position within this one photo is all that's needed to put
      // them back in reading order. One line per grid row, one cell per
      // line: core/estimateExtract.ts's extractEstimateItemsFromLines reads
      // the line text directly rather than resolving column positions.
      const sorted = [...lines].sort((a, b) => a.top - b.top)
      allRows.push(...sorted.map((l) => [l.text]))
    }
    return {
      id: `ocr-${Date.now()}`,
      name: 'Photo estimate',
      path: '',
      sheetName: 'Photo estimate',
      grid: allRows,
      startRow: 0
    }
  })

  ipcMain.handle(IPC.ocrMbMeasurementSheet, async (_e, dataUrls: string[]): Promise<SheetGrid> => {
    const buffers = dataUrls.map((u) => Buffer.from(u.split(',')[1] ?? '', 'base64'))
    const rows = await extractMbMeasurementSheet(buffers, (progress) => {
      mainWindow?.webContents.send(IPC.mbMeasurementProgress, progress)
    })
    const grid = mbMeasurementRowsToGrid(rows)
    return {
      id: `mb-measurement-${Date.now()}`,
      name: 'MB Measurement Sheet',
      path: '',
      sheetName: 'MB Measurement Sheet',
      grid,
      startRow: 0
    }
  })

  ipcMain.handle(IPC.openPath, async (_e, target: string): Promise<void> => {
    if (/^https?:\/\//i.test(target)) {
      await shell.openExternal(target)
      return
    }
    await shell.openPath(target)
  })

  ipcMain.handle(IPC.revealItem, async (_e, target: string): Promise<void> => {
    shell.showItemInFolder(target)
  })

  ipcMain.handle(IPC.defaultDir, async (): Promise<string> => {
    return app.getPath('downloads')
  })

  ipcMain.handle(IPC.getAppVersion, async (): Promise<string> => {
    return app.getVersion()
  })

  // Scrape a Telangana Government holiday-calendar page — which year it is
  // depends entirely on the URL the renderer passes in (see Dashboard.tsx:
  // the government publishes one page per year, e.g. .../calendar-2026/,
  // .../calendar-2027/ once that year's page goes up, so rather than this
  // app guessing/hardcoding the year, the user pastes in whichever year's
  // link they need). Cached per URL so switching the link to a new year
  // doesn't serve the old year's stale cache.
  ipcMain.handle(IPC.fetchCalendar, async (_e, url: string, force?: boolean) => {
    if (!/^https:\/\//i.test(url)) throw new Error('Calendar link must be a valid https:// URL.')
    // The year is only used for display/caching, not to pick the URL — pulled
    // from the link itself (falls back to the current year if the link has
    // no obvious one) so it always matches whatever page was actually fetched.
    const year = url.match(/(\d{4})/)?.[1] ?? String(new Date().getFullYear())
    const cacheFile = path.join(app.getPath('userData'), `calendar-cache-${year}.json`)

    // Serve the cached copy unless a refresh is explicitly requested.
    if (!force) {
      try {
        const cached = JSON.parse(fs.readFileSync(cacheFile, 'utf8')) as CalendarData
        if (cached && Array.isArray(cached.months) && cached.months.length > 0) {
          return cached
        }
      } catch {
        // no/invalid cache — fall through to network fetch
      }
    }

    const html = await new Promise<string>((resolve, reject) => {
      const req = https.get(
        url,
        {
          headers: {
            accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
            'user-agent':
              'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36'
          }
        },
        (res) => {
          if (res.statusCode && res.statusCode >= 400) {
            reject(new Error(`Calendar page returned HTTP ${res.statusCode}`))
            res.resume()
            return
          }
          let body = ''
          res.setEncoding('utf8')
          res.on('data', (c) => (body += c))
          res.on('end', () => resolve(body))
        }
      )
      req.on('error', reject)
      req.setTimeout(15000, () => req.destroy(new Error('Calendar request timed out')))
    })
    const parsed = parseCalendarHtml(html, year)
    try {
      fs.writeFileSync(cacheFile, JSON.stringify(parsed), 'utf8')
    } catch {
      // cache write is best-effort
    }
    return parsed
  })

  ipcMain.handle(IPC.importFromLink, async (_e, url: string): Promise<ExcelTable> => {
    return importTableFromGoogleLink(url)
  })

  ipcMain.handle(IPC.importAllSheetsFromLink, async (_e, url: string): Promise<SheetGrid[]> => {
    return importAllSheetsFromGoogleLink(url)
  })

  ipcMain.handle(IPC.login, async (_e, loginId: string, password: string, forceLogout?: boolean): Promise<LoginResult> => {
    let result: LoginResult
    try {
      result = await validateLogin(loginId, password)
      if (result.ok) rememberOfflineAuth(loginId, password)
    } catch (e) {
      // The credentials sheet is unreachable — almost always no internet.
      // Only let this through if THIS SAME loginId/password has already
      // been verified online on this device before; a different login ID,
      // or one that's never succeeded online here, still gets the original
      // network error surfaced (never a blanket offline bypass).
      if (checkOfflineAuth(loginId, password)) {
        result = { ok: true, offline: true }
      } else {
        throw e
      }
    }
    if (!result.ok) return result

    // Point the local cache at THIS login's own file before anything reads or
    // writes it, so we never load the previous user's state and never write
    // this user's over theirs.
    currentLoginId = firebaseSync.normalizeId(loginId)

    // forceLogout only ever reaches here after the password above has
    // already been re-verified — the login screen only offers it once a
    // normal attempt has already come back maxSessions, using the same
    // loginId/password the user just typed. It immediately ends every other
    // device's session, so it's never applied silently.
    const { claim, remoteState } = await firebaseSync.startSession(
      loginId,
      os.hostname(),
      (partial) => {
        if (mainWindow && !mainWindow.isDestroyed()) {
          mainWindow.webContents.send(IPC.remoteStateUpdate, partial)
        }
      },
      forceLogout
    )
    if (!claim.ok) return { ok: false, maxSessions: true }

    if (remoteState) {
      try {
        const file = stateFile()
        // A record entered on this device but not yet reflected in the cloud
        // (pushState is fire-and-forget and swallows failures — a network
        // hiccup, or the app quitting/logging out before that push finished)
        // must not vanish just because THIS login's remote pull raced ahead
        // of it. Union the per-item lists with whatever this login's own
        // on-disk cache already has, keeping any local-only record instead of
        // letting a stale-but-present cloud copy silently overwrite it.
        let merged: PersistedState = remoteState
        try {
          const existing = JSON.parse(fs.readFileSync(file, 'utf8')) as PersistedState
          const unionById = <T extends { id: string }>(remote: T[] | undefined, local: T[] | undefined): T[] => {
            const remoteList = remote ?? []
            const seen = new Set(remoteList.map((it) => it.id))
            return [...remoteList, ...(local ?? []).filter((it) => !seen.has(it.id))]
          }
          merged = {
            ...remoteState,
            todos: unionById(remoteState.todos, existing.todos),
            mbScrutiny: unionById(remoteState.mbScrutiny, existing.mbScrutiny),
            createdDocuments: unionById(remoteState.createdDocuments, existing.createdDocuments)
          }
        } catch {
          // No readable existing cache for this login (fresh machine/first
          // login here) — remote alone is all there is.
        }
        fs.writeFileSync(file, JSON.stringify(merged), 'utf8')
      } catch {
        // Non-fatal: local cache write is best-effort.
      }
    }
    return result
  })

  ipcMain.handle(IPC.logout, async (): Promise<void> => {
    await firebaseSync.endSession()
    // Forget which login's cache we were pointing at, so nothing can read or
    // write it until the next login re-establishes identity.
    currentLoginId = null
  })

  ipcMain.handle(IPC.listActiveSessions, async (): Promise<ActiveSessionInfo[]> => {
    if (!currentLoginId) return []
    const slots = await firebaseSync.listSessions(currentLoginId)
    const mySessionId = firebaseSync.currentSessionId()
    return slots.map((s) => ({ ...s, isThisDevice: s.sessionId === mySessionId }))
  })

  ipcMain.handle(IPC.logoutOtherSession, async (_e, sessionId: string): Promise<void> => {
    if (!currentLoginId) return
    // Never let this end THIS device's own session through this path — that
    // would desync the renderer's auth state (still "logged in" locally with
    // no session backing it). Use IPC.logout for that instead.
    if (sessionId === firebaseSync.currentSessionId()) return
    await firebaseSync.endOtherSession(currentLoginId, sessionId)
  })

  ipcMain.on(IPC.restartToUpdate, () => {
    restartToUpdate()
  })

  ipcMain.handle(IPC.checkForUpdates, async (): Promise<ManualCheckResult> => {
    return checkForUpdatesManually()
  })

  ipcMain.handle(
    IPC.exportTable,
    async (_e, table: ExcelTable, suggestedName: string): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: `Save ${suggestedName}`,
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null
      fs.writeFileSync(result.filePath, buildWorkbookBuffer(table))
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.exportScheduleA,
    async (_e, table: ExcelTable, suggestedName: string, meta?: ScheduleAMeta): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: `Save ${suggestedName}`,
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const items = rowsToScheduleAItems(table)
      const templatePath = scheduleATemplateFile()
      const buffer = templatePath
        ? await fillScheduleATemplate(fs.readFileSync(templatePath), items, meta)
        : buildScheduleAWorkbook(items, meta) // fallback if the bundled template is missing

      fs.writeFileSync(result.filePath, buffer)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.exportSeScheduleA,
    async (_e, table: ExcelTable, suggestedName: string, meta?: ScheduleAMeta): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: `Save ${suggestedName}`,
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const items = rowsToScheduleAItems(table)
      const templatePath = seScheduleATemplateFile()
      if (!templatePath) throw new Error('SE Schedule A format is missing from the app bundle.')
      const buffer = await fillSeScheduleATemplate(fs.readFileSync(templatePath), items, meta)

      fs.writeFileSync(result.filePath, buffer)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.exportEvaluationSheet,
    async (_e, input: EvaluationSheetInput, suggestedName: string): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: `Save ${suggestedName}`,
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null
      fs.writeFileSync(result.filePath, await buildEvaluationSheet(input))
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.exportEvaluationSheetBatch,
    async (_e, entries: { input: EvaluationSheetInput; suggestedName: string }[]): Promise<string[] | null> => {
      const result = await dialog.showOpenDialog(mainWindow!, {
        title: 'Choose a folder to save all evaluation sheets into',
        properties: ['openDirectory', 'createDirectory']
      })
      if (result.canceled || result.filePaths.length === 0) return null
      const dir = result.filePaths[0]

      const used = new Set<string>()
      const written: string[] = []
      for (const entry of entries) {
        const base = entry.suggestedName || 'Evaluation Sheet'
        let fileName = `${base}.xlsx`
        let n = 2
        while (used.has(fileName) || fs.existsSync(path.join(dir, fileName))) {
          fileName = `${base} (${n}).xlsx`
          n += 1
        }
        used.add(fileName)
        fs.writeFileSync(path.join(dir, fileName), await buildEvaluationSheet(entry.input))
        written.push(path.join(dir, fileName))
      }
      return written
    }
  )

  async function buildScheduleABuffer(table: ExcelTable, meta?: ScheduleAMeta, isSe?: boolean): Promise<Buffer> {
    const items = rowsToScheduleAItems(table)
    // A Zone-level (SE) office must never get the Executive Engineer / Circle
    // template — same split as the standalone "Save Schedule A" button
    // (IPC.exportScheduleA / exportSeScheduleA above); this bundle path had
    // been missing that check entirely, always using the EE template/
    // signature even when saving from an SE office's Agreement tab.
    if (isSe) {
      const sePath = seScheduleATemplateFile()
      if (!sePath) throw new Error('SE Schedule A format is missing from the app bundle.')
      return fillSeScheduleATemplate(fs.readFileSync(sePath), items, meta)
    }
    const templatePath = scheduleATemplateFile()
    return templatePath
      ? fillScheduleATemplate(fs.readFileSync(templatePath), items, meta)
      : buildScheduleAWorkbook(items, meta)
  }

  // Save every agreement-workspace document into ONE folder the user picks —
  // each in its requested format (docx as-is, pdf via LibreOffice, xlsx built
  // from the Schedule A table). A unique-name guard avoids clobbering.
  ipcMain.handle(
    IPC.exportAgreementBundle,
    async (_e, files: AgreementBundleFile[]): Promise<{ written: string[]; failed: string[] } | null> => {
      if (!files || files.length === 0) return null
      const result = await dialog.showOpenDialog(mainWindow!, {
        title: 'Choose a folder to save all agreement documents into',
        properties: ['openDirectory', 'createDirectory']
      })
      if (result.canceled || result.filePaths.length === 0) return null
      const dir = result.filePaths[0]

      const used = new Set<string>()
      const uniquePath = (base: string, ext: string): string => {
        let fileName = `${base}.${ext}`
        let n = 2
        while (used.has(fileName) || fs.existsSync(path.join(dir, fileName))) {
          fileName = `${base} (${n}).${ext}`
          n += 1
        }
        used.add(fileName)
        return path.join(dir, fileName)
      }

      const written: string[] = []
      const failed: string[] = []
      let doneCount = 0
      for (const f of files) {
        mainWindow?.webContents.send(IPC.agreementBundleProgress, { done: doneCount, total: files.length, name: f.name })
        try {
          if (f.format === 'xlsx') {
            if (!f.scheduleATable) continue
            const p = uniquePath(f.name, 'xlsx')
            fs.writeFileSync(p, await buildScheduleABuffer(f.scheduleATable, f.scheduleAMeta, f.scheduleAIsSe))
            written.push(p)
          } else {
            if (!f.docxBase64) continue
            const docxBuffer = sanitizeDocxForWord2007(Buffer.from(f.docxBase64, 'base64'))
            if (f.format === 'pdf') {
              const p = uniquePath(f.name, 'pdf')
              fs.writeFileSync(p, await convertDocxToPdf(docxBuffer))
              written.push(p)
            } else {
              const p = uniquePath(f.name, 'docx')
              fs.writeFileSync(p, docxBuffer)
              written.push(p)
            }
          }
        } catch (e) {
          // Best-effort per file: one bad convert (e.g. LibreOffice missing for a
          // pdf) shouldn't abort the whole bundle — but record it so the caller
          // can tell the user exactly which documents didn't make it (they used
          // to vanish silently).
          console.error(`exportAgreementBundle: failed to write "${f.name}" (${f.format})`, e)
          failed.push(`${f.name} (${f.format.toUpperCase()})`)
        } finally {
          doneCount += 1
        }
      }
      return written.length > 0 || failed.length > 0 ? { written, failed } : null
    }
  )

  async function buildBoqBuffer(table: ExcelTable, workName?: string): Promise<Buffer> {
    const rows = rowsToBoqData(table)
    const templatePath = boqTemplateFile()
    return templatePath
      ? fillBoqTemplate(fs.readFileSync(templatePath), rows, workName)
      : buildWorkbookBuffer(table) // fallback if the bundled template is missing
  }

  ipcMain.handle(
    IPC.exportBoq,
    async (_e, table: ExcelTable, suggestedName: string, workName?: string): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: `Save ${suggestedName}`,
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null

      fs.writeFileSync(result.filePath, await buildBoqBuffer(table, workName))
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.exportBoqBatch,
    async (
      _e,
      entries: { table: ExcelTable; suggestedName: string; workName?: string }[]
    ): Promise<string[] | null> => {
      const result = await dialog.showOpenDialog(mainWindow!, {
        title: 'Choose a folder to save all BOQs into',
        properties: ['openDirectory', 'createDirectory']
      })
      if (result.canceled || result.filePaths.length === 0) return null
      const dir = result.filePaths[0]

      const used = new Set<string>()
      const written: string[] = []
      for (const entry of entries) {
        const base = entry.suggestedName || 'BOQ'
        let fileName = `${base}.xlsx`
        let n = 2
        while (used.has(fileName) || fs.existsSync(path.join(dir, fileName))) {
          fileName = `${base} (${n}).xlsx`
          n += 1
        }
        used.add(fileName)
        fs.writeFileSync(path.join(dir, fileName), await buildBoqBuffer(entry.table, entry.workName))
        written.push(path.join(dir, fileName))
      }
      return written
    }
  )

  ipcMain.handle(
    IPC.pickWorkbookForSplit,
    async (): Promise<{ path: string; name: string; sheets: SheetPreview[] } | null> => {
      const pick = await dialog.showOpenDialog(mainWindow!, {
        title: 'Select an Excel workbook to split into separate sheets',
        filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }],
        properties: ['openFile']
      })
      if (pick.canceled || pick.filePaths.length === 0) return null
      const srcPath = pick.filePaths[0]
      // A lightweight per-sheet preview (name + a corner of cells) drives the
      // live sheet tiles; the split itself still works off the sheet names.
      const sheets = readSheetPreviews(srcPath)
      return { path: srcPath, name: path.basename(srcPath), sheets }
    }
  )

  ipcMain.handle(
    IPC.splitWorkbook,
    async (_e, srcPath: string, sheetNames: string[] | null): Promise<{ dir: string; files: string[] } | null> => {
      // Prompt for the destination folder. buttonLabel + a defaultPath in the
      // source workbook's own folder make it clear this dialog is asking *where
      // to save the split sheets*, not to re-pick a file.
      const folder = await dialog.showOpenDialog(mainWindow!, {
        title: 'Choose where to save the separated sheets',
        buttonLabel: 'Save sheets here',
        defaultPath: path.dirname(srcPath),
        properties: ['openDirectory', 'createDirectory']
      })
      if (folder.canceled || folder.filePaths.length === 0) return null
      const dir = folder.filePaths[0]

      // Run the split in an isolated child process (raised heap): a large
      // workbook that would OOM is failed gracefully here rather than crashing
      // the whole app.
      const files = await runSplitInWorker(srcPath, dir, sheetNames, (done, total, sheet) => {
        mainWindow?.webContents.send(IPC.splitProgress, { done, total, sheet })
      })
      return { dir, files }
    }
  )

  ipcMain.handle(IPC.pickPdfsForMerge, async (): Promise<{ path: string; name: string; pages: number }[] | null> => {
    const pick = await dialog.showOpenDialog(mainWindow!, {
      title: 'Select the PDFs to merge (in order)',
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
      properties: ['openFile', 'multiSelections']
    })
    if (pick.canceled || pick.filePaths.length === 0) return null
    const out: { path: string; name: string; pages: number }[] = []
    for (const p of pick.filePaths) {
      let pages = 0
      try {
        pages = await pdfPageCount(p)
      } catch {
        pages = 0
      }
      out.push({ path: p, name: path.basename(p), pages })
    }
    return out
  })

  ipcMain.handle(IPC.mergePdfs, async (_e, srcPaths: string[]): Promise<{ file: string } | null> => {
    if (!srcPaths || srcPaths.length === 0) return null
    const save = await dialog.showSaveDialog(mainWindow!, {
      title: 'Save the merged PDF',
      defaultPath: path.join(path.dirname(srcPaths[0]), 'Merged.pdf'),
      filters: [{ name: 'PDF', extensions: ['pdf'] }]
    })
    if (save.canceled || !save.filePath) return null
    const file = save.filePath.toLowerCase().endsWith('.pdf') ? save.filePath : `${save.filePath}.pdf`
    await mergePdfFiles(srcPaths, file)
    return { file }
  })

  ipcMain.handle(IPC.pickPdfForSplit, async (): Promise<{ path: string; name: string; pages: number } | null> => {
    const pick = await dialog.showOpenDialog(mainWindow!, {
      title: 'Select a PDF to separate',
      filters: [{ name: 'PDF', extensions: ['pdf'] }],
      properties: ['openFile']
    })
    if (pick.canceled || pick.filePaths.length === 0) return null
    const p = pick.filePaths[0]
    return { path: p, name: path.basename(p), pages: await pdfPageCount(p) }
  })

  ipcMain.handle(
    IPC.splitPdf,
    async (_e, srcPath: string, ranges: [number, number][] | null): Promise<{ dir: string; files: string[] } | null> => {
      const folder = await dialog.showOpenDialog(mainWindow!, {
        title: 'Choose where to save the separated PDFs',
        buttonLabel: 'Save PDFs here',
        defaultPath: path.dirname(srcPath),
        properties: ['openDirectory', 'createDirectory']
      })
      if (folder.canceled || folder.filePaths.length === 0) return null
      const dir = folder.filePaths[0]
      const files = await splitPdfFile(srcPath, dir, ranges)
      return { dir, files }
    }
  )

  // PDF workspace — the renderer builds the output PDF(s) with pdf-lib (the same
  // pages the user selected across the uploaded files) and hands the raw bytes
  // here just to be written to disk via the OS save/folder dialog.
  ipcMain.handle(IPC.savePdf, async (_e, bytes: Uint8Array, suggestedName: string): Promise<{ file: string } | null> => {
    const save = await dialog.showSaveDialog(mainWindow!, {
      title: 'Save the PDF',
      defaultPath: suggestedName || 'Document.pdf',
      filters: [{ name: 'PDF', extensions: ['pdf'] }]
    })
    if (save.canceled || !save.filePath) return null
    const file = save.filePath.toLowerCase().endsWith('.pdf') ? save.filePath : `${save.filePath}.pdf`
    fs.writeFileSync(file, Buffer.from(bytes))
    return { file }
  })

  ipcMain.handle(
    IPC.savePdfsToFolder,
    async (_e, files: { name: string; bytes: Uint8Array }[]): Promise<{ dir: string; files: string[] } | null> => {
      if (!files || files.length === 0) return null
      const folder = await dialog.showOpenDialog(mainWindow!, {
        title: 'Choose where to save the PDFs',
        buttonLabel: 'Save PDFs here',
        properties: ['openDirectory', 'createDirectory']
      })
      if (folder.canceled || folder.filePaths.length === 0) return null
      const dir = folder.filePaths[0]
      const written: string[] = []
      for (const f of files) {
        // Strip anything that can't be a file name so a source PDF's name can't
        // escape the chosen folder or break the write.
        const safe = f.name.replace(/[\\/:*?"<>|]/g, '_')
        const out = path.join(dir, safe.toLowerCase().endsWith('.pdf') ? safe : `${safe}.pdf`)
        fs.writeFileSync(out, Buffer.from(f.bytes))
        written.push(out)
      }
      return { dir, files: written }
    }
  )

  // GPS Photos tool — OCR the stamped overlay when a photo has no EXIF GPS.
  ipcMain.handle(IPC.ocrGpsOverlay, async (_e, imageBytes: Uint8Array): Promise<string[]> => {
    return ocrGpsOverlay(Buffer.from(imageBytes))
  })

  // Cement & Steel Rates tool — list circulars from the Public Health
  // department's Downloads page, and save one via its short-lived token.
  ipcMain.handle(IPC.fetchCementSteelRates, async (): Promise<CementSteelRate[]> => {
    return fetchCementSteelRates()
  })

  ipcMain.handle(
    IPC.downloadCementSteelRate,
    async (_e, token: string, suggestedFileName: string): Promise<string | null> => {
      const buffer = await downloadCementSteelRateBuffer(token)
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: `Save ${suggestedFileName}`,
        defaultPath: suggestedFileName
      })
      if (result.canceled || !result.filePath) return null
      fs.writeFileSync(result.filePath, buffer)
      return result.filePath
    }
  )

  // Photos/PDF → Word/Excel tool — OCR each page image (in order) into text
  // lines, kept in reading order and separated by a blank line between pages so
  // the renderer can show one editable block the user reviews before exporting.
  ipcMain.handle(IPC.ocrPhotosToLines, async (_e, dataUrls: string[]): Promise<string[]> => {
    const buffers = dataUrls.map((u) => Buffer.from(u.split(',')[1] ?? '', 'base64'))
    // OCR every page in parallel across the process pool (one per core) instead
    // of one at a time — the big speed-up for multi-page uploads.
    const perPage = await recognizeImages(buffers)
    const out: string[] = []
    perPage.forEach((lines, i) => {
      const sorted = [...lines].sort((a, b) => a.top - b.top)
      if (i > 0) out.push('') // blank line between pages
      out.push(...sorted.map((l) => l.text))
    })
    return out
  })

  ipcMain.handle(
    IPC.savePhotosAsWord,
    async (_e, text: string, suggestedName: string): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save as Word',
        defaultPath: `${suggestedName || 'Document'}.docx`,
        filters: [{ name: 'Word Document', extensions: ['docx'] }]
      })
      if (result.canceled || !result.filePath) return null
      const buffer = await convertHtmlToDocx(textToParagraphsHtml(text))
      fs.writeFileSync(result.filePath, buffer)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.savePhotosAsExcel,
    async (_e, text: string, suggestedName: string): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save as Excel',
        defaultPath: `${suggestedName || 'Spreadsheet'}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null
      fs.writeFileSync(result.filePath, buildPhotosWorkbook(text, suggestedName))
      return result.filePath
    }
  )

  // Photos/PDF → Word (keep layout): convert the ORIGINAL uploaded PDF(s) to a
  // .docx that preserves the layout (LibreOffice writer_pdf_import), merging
  // several PDFs in order — vs the OCR path that only extracts plain text.
  ipcMain.handle(
    IPC.savePdfAsWord,
    async (_e, pdfs: { name: string; bytes: Uint8Array }[], suggestedName: string): Promise<string | null> => {
      if (!pdfs || pdfs.length === 0) return null
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save as Word (keep layout)',
        defaultPath: `${suggestedName || 'Document'}.docx`,
        filters: [{ name: 'Word Document', extensions: ['docx'] }]
      })
      if (result.canceled || !result.filePath) return null
      const docxBuffers: Buffer[] = []
      for (const pdf of pdfs) docxBuffers.push(await convertPdfToDocx(Buffer.from(pdf.bytes)))
      const merged = docxBuffers.length === 1 ? docxBuffers[0] : mergeDocxBuffers(docxBuffers)
      // convertPdfToDocx shells out to LibreOffice's own PDF->docx conversion —
      // the same authoring tool responsible for every other bidi/child-order
      // corruption found in this codebase, so its output needs the same
      // defensive sanitizing before it reaches a real Word install.
      fs.writeFileSync(result.filePath, sanitizeDocxForWord2007(merged))
      return result.filePath
    }
  )

  // Photos/PDF → Word (offline reconstruction): the renderer rebuilds a text
  // PDF's layout (or plain OCR lines) into a doc-model of real paragraphs +
  // tables; this writes it as a directly-built, Word-valid .docx (core/docxBuilder
  // — NOT html-to-docx, whose output Microsoft Word refuses to open).
  ipcMain.handle(
    IPC.saveWordDoc,
    async (_e, blocks: DocBlock[], suggestedName: string): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save as Word',
        defaultPath: `${suggestedName || 'Document'}.docx`,
        filters: [{ name: 'Word Document', extensions: ['docx'] }]
      })
      if (result.canceled || !result.filePath) return null
      fs.writeFileSync(result.filePath, buildDocx(blocks))
      return result.filePath
    }
  )

  // Photos/PDF → Word/Excel (offline image-based reconstruction): OCR each page
  // and return the recognised lines WITH their bounding boxes, per page, so the
  // renderer/core can rebuild the table from where the text sits on the image
  // (handles photos, scans, and broken-font pages the text layer can't).
  ipcMain.handle(IPC.ocrPhotosToLayout, async (_e, dataUrls: string[]): Promise<OcrPage[]> => {
    const buffers = dataUrls.map((u) => Buffer.from(u.split(',')[1] ?? '', 'base64'))
    const perPage = await recognizeImages(buffers)
    return perPage.map((lines) => ({
      lines: lines
        .filter((l) => l.box && l.text.trim())
        .map((l) => ({ text: l.text, x: l.box![0], y: l.box![1], w: l.box![2], h: l.box![3] }))
    }))
  })

  ipcMain.handle(
    IPC.saveRowsAsExcel,
    async (_e, rows: string[][], suggestedName: string): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save as Excel',
        defaultPath: `${suggestedName || 'Spreadsheet'}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null
      fs.writeFileSync(result.filePath, buildWorkbookFromRows(rows, suggestedName))
      return result.filePath
    }
  )

  // Word workspace — a .docx has no addressable pages until it's laid out, so the
  // page-level preview is driven off a LibreOffice conversion to PDF. The renderer
  // then renders/selects/merges those pages exactly like the PDF workspace.
  ipcMain.handle(IPC.docxToPdf, async (_e, docxBytes: Uint8Array): Promise<Uint8Array> => {
    const pdf = await convertDocxToPdf(Buffer.from(docxBytes))
    return new Uint8Array(pdf)
  })

  ipcMain.handle(IPC.docxToPageImages, async (_e, docxBytes: Uint8Array): Promise<Uint8Array[]> => {
    const images = await docxToPageImages(Buffer.from(docxBytes))
    return images.map((buf) => new Uint8Array(buf))
  })

  // Whole-document merge: concatenate the .docx files (in order) into one and
  // save it via a dialog. Kept in Word format (unlike the page tool's PDF output).
  ipcMain.handle(IPC.mergeDocx, async (_e, docxBytesList: Uint8Array[]): Promise<{ file: string } | null> => {
    if (!docxBytesList || docxBytesList.length === 0) return null
    const merged = mergeDocxBuffers(docxBytesList.map((b) => Buffer.from(b)))
    const save = await dialog.showSaveDialog(mainWindow!, {
      title: 'Save the merged Word document',
      defaultPath: 'Merged.docx',
      filters: [{ name: 'Word Document', extensions: ['docx'] }]
    })
    if (save.canceled || !save.filePath) return null
    const file = save.filePath.toLowerCase().endsWith('.docx') ? save.filePath : `${save.filePath}.docx`
    fs.writeFileSync(file, merged)
    return { file }
  })

  // Split a .docx into one .docx per page (at its page breaks), each keeping full
  // formatting — pure XML surgery, so it works without LibreOffice (only the
  // page previews need it).
  ipcMain.handle(IPC.splitDocxSections, async (_e, docxBytes: Uint8Array): Promise<Uint8Array[]> => {
    return splitDocxByPageBreaks(Buffer.from(docxBytes)).map((b) => new Uint8Array(b))
  })

  ipcMain.handle(
    IPC.saveDocxsToFolder,
    async (_e, files: { name: string; bytes: Uint8Array }[]): Promise<{ dir: string; files: string[] } | null> => {
      if (!files || files.length === 0) return null
      const folder = await dialog.showOpenDialog(mainWindow!, {
        title: 'Choose where to save the Word files',
        buttonLabel: 'Save here',
        properties: ['openDirectory', 'createDirectory']
      })
      if (folder.canceled || folder.filePaths.length === 0) return null
      const dir = folder.filePaths[0]
      const written: string[] = []
      for (const f of files) {
        const safe = f.name.replace(/[\\/:*?"<>|]/g, '_')
        const out = path.join(dir, safe.toLowerCase().endsWith('.docx') ? safe : `${safe}.docx`)
        fs.writeFileSync(out, Buffer.from(f.bytes))
        written.push(out)
      }
      return { dir, files: written }
    }
  )

  ipcMain.handle(
    IPC.exportDeviation,
    async (
      _e,
      items: DeviationItem[],
      meta: DeviationMeta,
      suggestedName: string
    ): Promise<string | null> => {
      const templatePath = deviationTemplateFile()
      if (!templatePath) throw new Error('Deviation Statement template is missing from the app bundle.')

      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save Deviation Statement',
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const filled = await fillDeviationTemplate(fs.readFileSync(templatePath), items, meta)
      fs.writeFileSync(result.filePath, filled)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.exportDetailedEstimate,
    async (
      _e,
      items: EstimateWorkItem[],
      meta: DetailedEstimateMeta,
      suggestedName: string
    ): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save Estimate',
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const buffer = await buildDetailedEstimateWorkbook(items, meta)
      fs.writeFileSync(result.filePath, buffer)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.exportMaterialEstimate,
    async (_e, totals: MaterialTotals, meta: MaterialEstimateMeta, suggestedName: string): Promise<string | null> => {
      const templatePath = materialTemplateFile()
      if (!templatePath) throw new Error('Material Estimation template is missing from the app bundle.')

      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save Material Estimation',
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const buffer = await fillMaterialTemplate(fs.readFileSync(templatePath), totals, meta)
      fs.writeFileSync(result.filePath, buffer)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.searchTenders,
    async (_e, query: TenderQuery): Promise<TenderResult> => {
      const raw = await fetchTenders(query)
      let json: { aaData?: string[][]; iTotalRecords?: number; iTotalDisplayRecords?: number }
      try {
        json = JSON.parse(raw)
      } catch {
        throw new Error('The tender portal returned an unexpected response. Please try again.')
      }
      return {
        data: json.aaData ?? [],
        total: json.iTotalDisplayRecords ?? json.iTotalRecords ?? -1
      }
    }
  )

  ipcMain.handle(IPC.embedTexts, async (_e, texts: string[]): Promise<number[][]> => {
    return embedTexts(texts)
  })

  ipcMain.handle(IPC.listDocumentParagraphs, async (_e, docxBase64: string): Promise<string[]> => {
    return listParagraphs(Buffer.from(docxBase64, 'base64'))
  })

  ipcMain.handle(
    IPC.saveDocumentEdits,
    async (_e, docxBase64: string, edits: { index: number; text: string }[]): Promise<string> => {
      const updated = applyParagraphEdits(
        Buffer.from(docxBase64, 'base64'),
        edits.map((e) => ({ index: e.index, newText: e.text }))
      )
      return updated.toString('base64')
    }
  )

  ipcMain.handle(IPC.findPlaceholdersInDocument, async (_e, docxBase64: string): Promise<string[]> => {
    return findPlaceholdersInAllParts(Buffer.from(docxBase64, 'base64'))
  })

  ipcMain.handle(
    IPC.fillPlaceholdersInDocument,
    async (_e, docxBase64: string, resolved: PlaceholderMatch[], row: Record<string, string>): Promise<string> => {
      const filled = fillPlaceholdersInAllParts(Buffer.from(docxBase64, 'base64'), resolved, row)
      return filled.toString('base64')
    }
  )

  ipcMain.handle(
    IPC.bakeFixedPlaceholdersInDocument,
    async (_e, docxBase64: string, values: Record<string, string>): Promise<string> => {
      const baked = bakeFixedPlaceholdersInDocx(Buffer.from(docxBase64, 'base64'), values)
      return baked.toString('base64')
    }
  )

  ipcMain.handle(IPC.verifyDocuments, async (_e, items: VerifyDocItem[]): Promise<VerifyDocResult[]> => {
    return items.map((item) => {
      const text = extractAllText(Buffer.from(item.docxBase64, 'base64'))
      const checks = [verifyPlaceholderCoverage(text, item.values, item.requiredLabels)]
      if (item.amounts) checks.push(verifyAmountMath({ docText: text, ...item.amounts }))
      if (item.corporation) checks.push(verifyCorporationWording(text, item.corporation.expected, item.corporation.all))
      if (item.reserved != null) checks.push(verifyReservedTag(text, item.reserved))
      const result = combineVerify(...checks)
      return { name: item.name, ok: result.ok, issues: result.issues }
    })
  })

  ipcMain.handle(
    IPC.exportCreatedDocument,
    async (
      _e,
      docxBase64: string,
      suggestedName: string,
      formats: ('docx' | 'pdf')[]
    ): Promise<{ file: string; format: 'docx' | 'pdf' }[] | null> => {
      if (formats.length === 0) return null
      // Rewrite any Word-2007-incompatible border elements before the file
      // leaves the app — the single choke point every exported .docx and PDF
      // passes through, so every document type is covered in one place.
      const docxBuffer = sanitizeDocxForWord2007(Buffer.from(docxBase64, 'base64'))
      const written: { file: string; format: 'docx' | 'pdf' }[] = []

      if (formats.includes('docx')) {
        const result = await dialog.showSaveDialog(mainWindow!, {
          title: `Save ${suggestedName}`,
          defaultPath: `${suggestedName}.docx`,
          filters: [{ name: 'Word Document', extensions: ['docx'] }]
        })
        if (!result.canceled && result.filePath) {
          fs.writeFileSync(result.filePath, docxBuffer)
          written.push({ file: result.filePath, format: 'docx' })
        }
      }

      if (formats.includes('pdf')) {
        const result = await dialog.showSaveDialog(mainWindow!, {
          title: `Save ${suggestedName}`,
          defaultPath: `${suggestedName}.pdf`,
          filters: [{ name: 'PDF Document', extensions: ['pdf'] }]
        })
        if (!result.canceled && result.filePath) {
          const pdfBuffer = await convertDocxToPdf(docxBuffer)
          fs.writeFileSync(result.filePath, pdfBuffer)
          written.push({ file: result.filePath, format: 'pdf' })
        }
      }

      return written.length > 0 ? written : null
    }
  )

  let printWindow: BrowserWindow | null = null

  // Skip any in-app preview and go straight to the OS print dialog. Fed
  // docx-preview's own rendered HTML (captured by the renderer after it
  // renders the filled .docx) rather than raw pasted content — this avoids
  // needing LibreOffice for printing (only creating a document and PDF
  // export need it).
  ipcMain.handle(IPC.printCreatedDocument, async (_e, renderedHtml: string): Promise<void> => {
    // docx-preview output already renders a full A4 page (section.docx) with the
    // document's own 1-inch margins baked in, so it must print at zero page
    // margin — otherwise the browser's default ~0.4-inch print margin makes the
    // page wider than the printable area and clips the right edge. Plain-HTML
    // documents (Note Submitted, Schedule A) carry no page margin, so they get
    // a printed page margin instead. Detected by docx-preview's wrapper class.
    const isDocx = /docx-wrapper/.test(renderedHtml)
    const printStyle = isDocx
      ? `@page { size: A4; margin: 0 }
         html, body { margin: 0; padding: 0; background: #fff }
         .docx-wrapper { background: #fff !important; padding: 0 !important; margin: 0 !important }
         .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 auto !important }`
      : `@page { size: A4; margin: 0 }
         html, body { margin: 0; background: #fff }
         body { padding: 14mm 16mm }`
    const head = `<head><meta charset="utf-8"><style>${printStyle}</style></head>`
    const body = /<html/i.test(renderedHtml)
      ? renderedHtml
      : `<!DOCTYPE html><html>${head}<body>${renderedHtml}</body></html>`
    const file = path.join(app.getPath('temp'), `docugen-print-${Date.now()}.html`)
    fs.writeFileSync(file, body, 'utf8')

    if (printWindow && !printWindow.isDestroyed()) printWindow.close()
    // The OS print dialog (a sheet on macOS) needs a real, visible parent
    // window to attach to — a `show: false` window let the print call
    // silently go nowhere, with no dialog and no error, which read as
    // "print does nothing".
    printWindow = new BrowserWindow({
      width: 900,
      height: 1000,
      title: 'Print',
      webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false }
    })
    const win = printWindow
    win.on('closed', () => {
      if (printWindow === win) printWindow = null
      fs.rm(file, { force: true }, () => {})
    })
    await win.loadFile(file)
    win.show()
    win.focus()
    // Deliberately not closing the window from the print() callback: on
    // this Electron/macOS combination that callback can fire as soon as the
    // print sheet appears, not once the user has actually finished with it
    // — closing the window right then yanked the sheet away with it,
    // which looked like the dialog "opening and closing immediately". The
    // window is left open for the user to dismiss once they're done.
    win.webContents.print({ margins: { marginType: 'none' } }, (success, failureReason) => {
      if (!success && failureReason !== 'cancelled') {
        console.error('Print failed:', failureReason)
      }
    })
  })

  ipcMain.handle(IPC.noteSubmittedDocx, async (_e, html: string): Promise<string> => {
    const docxBuffer = await convertHtmlToDocx(html)
    return docxBuffer.toString('base64')
  })

  const bundledResourceFile = (fileName: string) =>
    [
      path.join(process.resourcesPath, fileName),
      path.join(app.getAppPath(), 'resources', fileName),
      path.join(app.getAppPath(), '..', 'resources', fileName)
    ].find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })

  ipcMain.handle(IPC.intimationTemplate, async (_e, variantId?: string): Promise<string> => {
    const templatePath = bundledResourceFile(intimationTemplateFileName(variantId))
    if (!templatePath) throw new Error('Intimation format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.workOrderTemplate, async (_e, variantId?: string): Promise<string> => {
    const templatePath = bundledResourceFile(workOrderTemplateFileName(variantId))
    if (!templatePath) throw new Error('Work Order format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.agreementTemplate, async (_e, variantId?: string): Promise<string> => {
    const templatePath = bundledResourceFile(agreementTemplateFileName(variantId))
    if (!templatePath) throw new Error('Agreement format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.fileBackerTemplate, async (_e, variantId?: string): Promise<string> => {
    const templatePath = bundledResourceFile(fileBackerTemplateFileName(variantId))
    if (!templatePath) throw new Error('File Backer format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.qccIntimationTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('qcc-intimation-template.docx')
    if (!templatePath) throw new Error('QCC Intimation format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.forwardingSlipTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('forwarding-slip-template.docx')
    if (!templatePath) throw new Error('Forwarding Slip format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.civilTenderTemplate, async (_e, variantId?: string): Promise<string> => {
    const templatePath = bundledResourceFile(civilTenderTemplateFileName(variantId))
    if (!templatePath) throw new Error('Civil Tender Document format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.seAgreementBondTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('se-agreement-bond-template.docx')
    if (!templatePath) throw new Error('SE Agreement Bond format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.zonalWorkOrderTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('zonal-work-order-template.docx')
    if (!templatePath) throw new Error('SE Work Order format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.zonalConcludingAgreementTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('zonal-concluding-agreement-template.docx')
    if (!templatePath) throw new Error('SE Concluding Agreement format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.zonalMemoEeTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('zonal-memo-ee-template.docx')
    if (!templatePath) throw new Error('SE Memo to EE format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.seAgreementNoteTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('se-agreement-note-template.docx')
    if (!templatePath) throw new Error('SE Agreement Put-up Note format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.contractDeedTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('contract-deed-template.docx')
    if (!templatePath) throw new Error('Contract Deed format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.loaSeTemplate, async (_e, reserved: boolean, corporation?: string): Promise<string> => {
    const isMmc = corporation === 'MMC'
    const file = isMmc
      ? reserved
        ? 'loa-se-mmc-reserved-template.docx'
        : 'loa-se-mmc-template.docx'
      : reserved
        ? 'loa-se-reserved-template.docx'
        : 'loa-se-template.docx'
    const templatePath = bundledResourceFile(file)
    if (!templatePath) throw new Error('Superintending Engineer LOA format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.tsNoteTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('ts-note-template.docx')
    if (!templatePath) throw new Error('TS Note format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.eligibilityCriteriaTemplate, async (): Promise<string> => {
    const templatePath = bundledResourceFile('eligibility-criteria-template.docx')
    if (!templatePath) throw new Error('Eligibility Criteria format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  ipcMain.handle(IPC.issueNoticeTemplate, async (_e, ee: boolean): Promise<string> => {
    const file = ee ? 'issue-notice-ee-template.docx' : 'issue-notice-template.docx'
    const templatePath = bundledResourceFile(file)
    if (!templatePath) throw new Error('Notice format is missing from the app bundle.')
    return fs.readFileSync(templatePath).toString('base64')
  })

  // Per-login local cache: state-<loginId>.json once someone has logged in, so
  // two logins on the same computer keep entirely separate on-disk workspaces
  // and one can never read the other's. Falls back to the legacy shared
  // state.json only when no login has happened yet (dev, login screen skipped).
  const stateFile = () =>
    path.join(app.getPath('userData'), currentLoginId ? `state-${currentLoginId}.json` : 'state.json')

  const seedStateFile = () => {
    // Bundled default state (works database + default Issue Document set) used on first run.
    const candidates = [
      path.join(process.resourcesPath, 'seed-state.json'),
      path.join(app.getAppPath(), 'resources', 'seed-state.json'),
      path.join(app.getAppPath(), '..', 'resources', 'seed-state.json')
    ]
    return candidates.find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })
  }

  const scheduleATemplateFile = () => {
    const candidates = [
      path.join(process.resourcesPath, 'schedule-a-template.xlsx'),
      path.join(app.getAppPath(), 'resources', 'schedule-a-template.xlsx'),
      path.join(app.getAppPath(), '..', 'resources', 'schedule-a-template.xlsx')
    ]
    return candidates.find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })
  }

  const seScheduleATemplateFile = () => bundledResourceFile('se-schedule-a-template.xlsx')

  const boqTemplateFile = () => {
    const candidates = [
      path.join(process.resourcesPath, 'boq-template.xlsx'),
      path.join(app.getAppPath(), 'resources', 'boq-template.xlsx'),
      path.join(app.getAppPath(), '..', 'resources', 'boq-template.xlsx')
    ]
    return candidates.find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })
  }

  const deviationTemplateFile = () => {
    const candidates = [
      path.join(process.resourcesPath, 'deviation-template.xlsx'),
      path.join(app.getAppPath(), 'resources', 'deviation-template.xlsx'),
      path.join(app.getAppPath(), '..', 'resources', 'deviation-template.xlsx')
    ]
    return candidates.find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })
  }

  const tenderNoticeTemplateFile = () => {
    const candidates = [
      path.join(process.resourcesPath, 'tender-notice-template.docx'),
      path.join(app.getAppPath(), 'resources', 'tender-notice-template.docx'),
      path.join(app.getAppPath(), '..', 'resources', 'tender-notice-template.docx')
    ]
    return candidates.find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })
  }

  const materialTemplateFile = () => {
    const candidates = [
      path.join(process.resourcesPath, 'material-estimation-template.xlsx'),
      path.join(app.getAppPath(), 'resources', 'material-estimation-template.xlsx'),
      path.join(app.getAppPath(), '..', 'resources', 'material-estimation-template.xlsx')
    ]
    return candidates.find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })
  }

  const bidDocumentTemplateFile = () => {
    const candidates = [
      path.join(process.resourcesPath, 'bid-document-template.docx'),
      path.join(app.getAppPath(), 'resources', 'bid-document-template.docx'),
      path.join(app.getAppPath(), '..', 'resources', 'bid-document-template.docx')
    ]
    return candidates.find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })
  }

  const bidDocumentSeTemplateFile = () => {
    const candidates = [
      path.join(process.resourcesPath, 'bid-document-se-template.docx'),
      path.join(app.getAppPath(), 'resources', 'bid-document-se-template.docx'),
      path.join(app.getAppPath(), '..', 'resources', 'bid-document-se-template.docx')
    ]
    return candidates.find((p) => {
      try {
        fs.accessSync(p)
        return true
      } catch {
        return false
      }
    })
  }

  // A Zone-level (SE) office has no Circle of its own — same "seMode" test
  // used throughout the app (GiveIntimationTab, WorkOrderAgreementTab,
  // TenderNoticeButton's canComposeNit) — so the SE Bid Document template
  // replaces the EE one for that work, rather than the two coexisting.
  const isSeWork = (work: BidDocumentInput['work']) => !!work.zone?.trim() && !work.circle?.trim()

  /** Picks the right bundled template + fill function for a work's office type, throwing a clear error if that template isn't in the app bundle. */
  const resolveBidDocumentFiller = (work: BidDocumentInput['work']) => {
    if (isSeWork(work)) {
      const templatePath = bidDocumentSeTemplateFile()
      if (!templatePath) throw new Error('SE Bid Document template is missing from the app bundle.')
      return { templatePath, fill: fillSeBidDocument }
    }
    const templatePath = bidDocumentTemplateFile()
    if (!templatePath) throw new Error('Bid Document template is missing from the app bundle.')
    return { templatePath, fill: fillBidDocument }
  }

  // Make an arbitrary label safe to use as a single file name. Win Codes and
  // work names can contain path separators (e.g. "16/DB/EE/…") and other
  // reserved characters; left unescaped these make path.join build a
  // non-existent sub-directory and fs.writeFileSync throw ENOENT — which, in
  // the batch loop, aborts after the first file so only one document saves.
  const sanitizeFileName = (name: string): string => {
    const cleaned = name
      .replace(/[/\\:*?"<>|\x00-\x1f]/g, '-') // path separators + OS-reserved chars
      .replace(/\s+/g, ' ')
      .replace(/[.\s]+$/g, '') // Windows disallows trailing dots/spaces
      .trim()
    return cleaned || 'Bid Document'
  }

  ipcMain.handle(
    IPC.generateTenderNotice,
    async (_e, input: TenderNoticeInput, suggestedName?: string): Promise<string | null> => {
      const templatePath = tenderNoticeTemplateFile()
      if (!templatePath) throw new Error('Tender Notice template is missing from the app bundle.')

      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save Tender Notice',
        defaultPath: `${suggestedName || 'Tender Notice'}.docx`,
        filters: [{ name: 'Word Document', extensions: ['docx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const filled = sanitizeDocxForWord2007(fillTenderNotice(fs.readFileSync(templatePath), input))
      fs.writeFileSync(result.filePath, filled)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.previewTenderNotice,
    async (_e, input: TenderNoticeInput): Promise<string> => {
      const templatePath = tenderNoticeTemplateFile()
      if (!templatePath) throw new Error('Tender Notice template is missing from the app bundle.')
      const filled = fillTenderNotice(fs.readFileSync(templatePath), input)
      return filled.toString('base64')
    }
  )

  ipcMain.handle(
    IPC.generateBidDocument,
    async (_e, input: BidDocumentInput, suggestedName: string): Promise<string | null> => {
      const { templatePath, fill } = resolveBidDocumentFiller(input.work)

      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save Bid Document',
        defaultPath: `${sanitizeFileName(suggestedName)}.docx`,
        filters: [{ name: 'Word Document', extensions: ['docx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const filled = sanitizeDocxForWord2007(fill(fs.readFileSync(templatePath), input))
      fs.writeFileSync(result.filePath, filled)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.generateBidDocumentBatch,
    async (
      _e,
      entries: { input: BidDocumentInput; suggestedName: string }[]
    ): Promise<string[] | null> => {
      const result = await dialog.showOpenDialog(mainWindow!, {
        title: 'Choose a folder to save all Bid Documents into',
        properties: ['openDirectory', 'createDirectory']
      })
      if (result.canceled || result.filePaths.length === 0) return null
      const dir = result.filePaths[0]

      // A single tender notice's works all share one office, so every entry
      // resolves to the same template — but each is still resolved on its
      // own (cheap: just a file-existence check) rather than assumed, so a
      // mixed batch (however unlikely) still gets the right template per row.
      const templateCache = new Map<string, Buffer>()
      const used = new Set<string>()
      const written: string[] = []
      for (let i = 0; i < entries.length; i++) {
        const entry = entries[i]
        const { templatePath, fill } = resolveBidDocumentFiller(entry.input.work)
        let templateBuffer = templateCache.get(templatePath)
        if (!templateBuffer) {
          templateBuffer = fs.readFileSync(templatePath)
          templateCache.set(templatePath, templateBuffer)
        }
        const base = sanitizeFileName(entry.suggestedName || 'Bid Document')
        let fileName = `${base}.docx`
        let n = 2
        while (used.has(fileName) || fs.existsSync(path.join(dir, fileName))) {
          fileName = `${base} (${n}).docx`
          n += 1
        }
        used.add(fileName)
        fs.writeFileSync(path.join(dir, fileName), sanitizeDocxForWord2007(fill(templateBuffer, entry.input)))
        written.push(path.join(dir, fileName))
        mainWindow?.webContents.send(IPC.bidDocumentBatchProgress, { done: i + 1, total: entries.length })
        // Filling + sanitizing a docx is CPU-bound and synchronous — looping
        // straight through a big batch with no yield hogs the main process's
        // single thread for the whole batch, freezing every window's IPC and
        // repaints along with it. Yielding once per document (not per-line
        // inside the fill itself) lets the event loop breathe between
        // documents without materially slowing the batch down.
        await new Promise((resolve) => setImmediate(resolve))
      }
      return written
    }
  )

  ipcMain.handle(
    IPC.previewBidDocument,
    async (_e, input: BidDocumentInput): Promise<string> => {
      const { templatePath, fill } = resolveBidDocumentFiller(input.work)
      const filled = fill(fs.readFileSync(templatePath), input)
      return filled.toString('base64')
    }
  )

  ipcMain.handle(
    IPC.generateTechnicalSanction,
    async (
      _e,
      estimatePath: string,
      sheetName: string,
      edits: CellEdit[],
      suggestedName: string,
      rateAnalysisRows?: (string | number)[][]
    ): Promise<string | null> => {
      const result = await dialog.showSaveDialog(mainWindow!, {
        title: `Save ${suggestedName}`,
        defaultPath: `${suggestedName}.xlsx`,
        filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const buffer = await applyTechnicalSanctionEdits(
        fs.readFileSync(estimatePath),
        sheetName,
        edits,
        rateAnalysisRows ?? []
      )
      fs.writeFileSync(result.filePath, buffer)
      return result.filePath
    }
  )

  // Pre-OOXML-refactor documents were saved with an `html` string instead of
  // a real `docx` buffer — converted once here (via the same html-to-docx
  // path exportCreatedDocument used to use) so an upgrading user's existing
  // documents (and the bundled seed-state.json's own 8 examples) keep
  // working rather than silently losing their content. Best-effort per
  // document: one bad conversion doesn't take the rest down with it.
  interface LegacyCreatedDocument {
    id: string
    name: string
    html?: string
    docx?: string
    createdDate: string
  }

  async function migrateCreatedDocuments(state: PersistedState): Promise<PersistedState> {
    const docs = (state.createdDocuments ?? []) as unknown as LegacyCreatedDocument[]
    if (docs.every((d) => d.docx)) return state
    const migrated = await Promise.all(
      docs.map(async (d) => {
        if (d.docx || !d.html) return d
        try {
          const docxBuffer = await convertHtmlToDocx(d.html)
          return { id: d.id, name: d.name, docx: docxBuffer.toString('base64'), createdDate: d.createdDate }
        } catch (e) {
          console.error(`Failed to migrate document "${d.name}" from html to docx`, e)
          return d
        }
      })
    )
    return { ...state, createdDocuments: migrated as unknown as PersistedState['createdDocuments'] }
  }

  // Standard documents bundled with the app that every workspace should carry.
  // Unlike the first-run seed (which only fills a brand-new install), these are
  // merged into an *existing* state too — once — the first time it's opened by a
  // build whose CURRENT_DEFAULT_DOC_VERSION exceeds the state's own
  // seededDocVersion. Keyed by a stable id so a document already present isn't
  // duplicated, and version-gated so one the user later deletes is never
  // re-added.
  // v11: drop the 8 legacy example documents (LEGACY_SEED_DOC_IDS) that early
  // builds seeded — they lingered on Issue Documents because injection only ever
  // added, never removed. Bumping the gate re-runs injection once on every
  // existing workspace so the prune below reaches installs already at v10.
  // v12: add the two QCC (Quality Control Cell) letters — the Intimation (before
  // work starts) and Completion (after work is done), both circle-scoped.
  // v13: fixed the QCC detail-table layout (fixed column grid — the value column
  // was starved and the work name wrapped into a giant cell); the bump re-runs
  // injection so the corrected template replaces the one already in a workspace.
  // v14: moved the QCC "To" recipient block (Quality Control Division) to the
  // right of the page (fixed From/To grid) — another in-place template refresh.
  // v15: dropped the "20" century-stub from the QCC date fields (Lr.No + Ref Dt).
  // v16: add the 3rd-party QC and 4th-party Intimation letters (circle-scoped).
  // v17: Public Participation Log Book header on one line — Division = "{{circle}}
  // circle", Sl. No. = {{cno}}.
  // v18: Action Taken Report — removed the empty padding rows that left a big gap
  // under Name of the Work; 4th-party letterhead uses "O/o Executive Engineer".
  // v19: Completion Report — Technical Sanction No & Date is one {{TS No and Date}}
  // placeholder so the " dt. " separator only shows when a value is present (a
  // no-work preview no longer prints a stray "dt.").
  // v20: Public Participation Log Book, Action Taken Report and Completion Report
  // are circle-scoped (EE office only) — the SE (zonal) office's Issue Documents
  // tab should show only its EOT proposal, not these EE-oriented reports.
  const CURRENT_DEFAULT_DOC_VERSION = 20
  const DEFAULT_DOCUMENTS: { id: string; name: string; file: string; officeScope?: 'zonal' | 'circle' }[] = [
    { id: 'doc_public_participation', name: 'Public Participation Log Book', file: 'public-participation-book-template.docx', officeScope: 'circle' },
    { id: 'doc_action_taken_report', name: 'Action Taken Report', file: 'action-taken-report-template.docx', officeScope: 'circle' },
    { id: 'doc_completion_report', name: 'Completion Report', file: 'completion-report-template.docx', officeScope: 'circle' },
    // The Superintending Engineer's EOT proposal belongs to the zone (SE) office;
    // the Executive Engineer's variant to a circle (EE) office. Scoped so each
    // shows only where it applies — see CreatedDocument.officeScope.
    { id: 'doc_eot_se', name: 'EOT Proposal (SE Office)', file: 'eot-se-template.docx', officeScope: 'zonal' },
    { id: 'doc_eot_ee', name: 'EOT Proposal (EE Office)', file: 'eot-ee-template.docx', officeScope: 'circle' },
    // Bill-stage forwarding notes at the Executive Engineer (circle) office: the
    // Dy.EE's covering letter forwarding the AE's bill, and the office note that
    // tabulates the bill for approval. Both circle-scoped on the Issue Documents
    // tab; the Tools tab shows them (and every doc) regardless of office.
    { id: 'doc_dy_ee_forwarding_note', name: 'Dy. EE Forwarding Note', file: 'dy-ee-forwarding-note-template.docx', officeScope: 'circle' },
    { id: 'doc_bill_forwarding_note', name: 'Bill Forwarding Note', file: 'bill-forwarding-note-template.docx', officeScope: 'circle' },
    // QCC (Quality Control Cell) letters to the Quality Control Division: the
    // Dy.EE's Intimation before a work starts, and the EE's Completion letter
    // once it's done. Both circle-scoped (EE office). The Intimation is also
    // offered on the Work Order/Agreement tab — same bundled template.
    { id: 'doc_qcc_intimation', name: 'QCC Intimation', file: 'qcc-intimation-template.docx', officeScope: 'circle' },
    { id: 'doc_qcc_completion', name: 'QCC Completion', file: 'qcc-completion-template.docx', officeScope: 'circle' },
    // EE letters to the 3rd-party (QC college) and 4th-party (testing lab). The
    // party agency (name/address/phone) changes yearly, so it's typed in at
    // issue time on the Issue Documents tab (see PrintDocumentTab's party fields),
    // not drawn from the Works List. Both circle-scoped.
    { id: 'doc_3rd_party_qc', name: '3rd Party QC', file: 'qcc-3rd-party-template.docx', officeScope: 'circle' },
    { id: 'doc_4th_party_intimation', name: '4th Party Intimation', file: 'qcc-4th-party-template.docx', officeScope: 'circle' }
  ]

  function injectDefaultDocuments(state: PersistedState): PersistedState {
    if ((state.seededDocVersion ?? 0) >= CURRENT_DEFAULT_DOC_VERSION) return state
    // Remove the superseded legacy example documents before (re)adding defaults.
    const docs = pruneLegacyDocuments([...(state.createdDocuments ?? [])])
    const createdDate = new Date().toISOString().slice(0, 10)
    for (const def of DEFAULT_DOCUMENTS) {
      const templatePath = bundledResourceFile(def.file)
      if (!templatePath) continue // best-effort: a missing bundle file doesn't block loading
      const docx = fs.readFileSync(templatePath).toString('base64')
      const existing = docs.findIndex((d) => d.id === def.id)
      if (existing >= 0) {
        // Already present: refresh the bundled template's content (and name /
        // scope) in place — a version bump also ships template fixes to an
        // existing workspace — while keeping the user's own ordering. Documents
        // can't be edited or deleted in-app, so nothing hand-changed is lost.
        docs[existing] = {
          ...docs[existing],
          name: def.name,
          docx,
          ...(def.officeScope ? { officeScope: def.officeScope } : {})
        }
      } else {
        docs.push({
          id: def.id,
          name: def.name,
          docx,
          createdDate,
          ...(def.officeScope ? { officeScope: def.officeScope } : {})
        })
      }
    }
    return { ...state, createdDocuments: docs, seededDocVersion: CURRENT_DEFAULT_DOC_VERSION }
  }

  ipcMain.handle(IPC.loadState, async (): Promise<PersistedState | null> => {
    let state: PersistedState | null = null
    try {
      state = JSON.parse(fs.readFileSync(stateFile(), 'utf8')) as PersistedState
    } catch {
      // First run (no saved state): seed with the bundled works database + default documents.
      try {
        const seed = seedStateFile()
        if (seed) state = JSON.parse(fs.readFileSync(seed, 'utf8')) as PersistedState
      } catch {
        // fall through
      }
    }
    return state && injectDefaultDocuments(await migrateCreatedDocuments(state))
  })

  // Content of the last state.json written, keyed by its path — so a debounced
  // autosave that fires with unchanged content (e.g. after applying a remote
  // echo) doesn't re-serialize and rewrite ~1MB to disk for nothing.
  let lastWrittenJson: string | null = null
  let lastWrittenPath: string | null = null
  ipcMain.handle(IPC.saveState, async (_e, state: PersistedState, skipCloud?: boolean): Promise<void> => {
    const file = stateFile()
    const json = JSON.stringify(state)
    if (json !== lastWrittenJson || file !== lastWrittenPath) {
      try {
        // Async write (not writeFileSync): a ~1MB serialize+write of the whole
        // workspace on every edit would otherwise block the main thread — the
        // one that services IPC and window events — and show up as UI lag.
        await fs.promises.writeFile(file, json, 'utf8')
        lastWrittenJson = json
        lastWrittenPath = file
      } catch {
        // Non-fatal: persistence is best-effort.
      }
    }
    // Don't block the local save on network latency — the cloud push is
    // itself best-effort and swallows its own errors. When this save merely
    // reflects a change we just received FROM the cloud (skipCloud), pushing it
    // straight back would echo it to the other sessions, which echo it back
    // again — a write storm in which a slightly-stale copy can clobber a task
    // another session just added. So persist locally but don't re-push.
    if (!skipCloud) void firebaseSync.pushState(state)
  })
}

// ── Tender portal bridge ─────────────────────────────────────────────
// The Telangana e-procurement site requires a real browser session (JSESSIONID)
// and blocks raw HTTP clients with a WAF. We host a hidden BrowserWindow on the
// site's own origin and run the JSON fetch inside it — same-origin, real cookies,
// no CORS — exactly like the portal's own AJAX.
const TENDER_ORIGIN = 'https://tender.telangana.gov.in'
const TENDER_ROOT = `${TENDER_ORIGIN}/`
const TENDER_HOME = `${TENDER_ORIGIN}/TenderDetailsHome.html`
let tenderWindow: BrowserWindow | null = null

/**
 * Create the hidden bridge window on a *fresh* in-memory partition. We avoid a
 * persistent partition on purpose: a persistent JSESSIONID that expires between
 * runs just redirects to the timeout page forever, and the on-disk partition
 * can also lock when two processes touch it. A unique in-memory partition name
 * guarantees a clean session every time we (re)create the window.
 */
function makeTenderWindow(): BrowserWindow {
  const win = new BrowserWindow({
    show: false,
    webPreferences: {
      partition: `tender-${Date.now()}`,
      contextIsolation: true,
      nodeIntegration: false
    }
  })
  win.on('closed', () => {
    if (tenderWindow === win) tenderWindow = null
  })
  return win
}

function getTenderWindow(): BrowserWindow {
  if (!tenderWindow || tenderWindow.isDestroyed()) tenderWindow = makeTenderWindow()
  return tenderWindow
}

/** Tear down the current bridge window so the next call gets a clean session. */
function recreateTenderWindow(): BrowserWindow {
  if (tenderWindow && !tenderWindow.isDestroyed()) tenderWindow.destroy()
  tenderWindow = makeTenderWindow()
  return tenderWindow
}

/**
 * Establish a portal session. The site root reliably issues a fresh JSESSIONID
 * and stays on https; loading the tender home then sets up the listing context
 * (it stays on TenderDetailsHome.html once a session exists).
 */
async function seatTenderSession(win: BrowserWindow): Promise<void> {
  await win.loadURL(TENDER_ROOT).catch(() => {})
  await win.loadURL(TENDER_HOME).catch(() => {})
}

/**
 * The JSON endpoint only responds with data when it receives the *complete*
 * DataTables querystring the portal's own page sends — a partial set returns an
 * HTML error page. We reproduce the full param set here and vary only paging,
 * listing type, and the cache-buster.
 */
function buildTenderUrl(query: TenderQuery): string {
  const params: Array<[string, string]> = [
    ['nTenderID', (query.search ?? '').trim()],
    ['nDepartmentID', '0'],
    ['subDeptId', ''],
    ['ddlDistrict', ''],
    ['ddlMandal', ''],
    ['biddingType', ''],
    ['sProcurementType', ''],
    ['mECVValue1', ''],
    ['mECVValue2', ''],
    ['dtBidClosingselect', ''],
    ['dtBidClosing1', ''],
    ['dtBidClosing2', ''],
    ['dtTenderOpening1', ''],
    ['dtTenderOpening2', ''],
    ['hdnSearch4', ''],
    ['hdnSearch', ''],
    ['hdncorrigendumsDetails', ''],
    ['hdncorrigendumsDetails1', ''],
    ['hdnnoSearch', '1'],
    ['hdncorrigendumsDetails2', ''],
    ['hdnPreviousPage', ''],
    ['hdnIndentID', ''],
    ['hdnTenderCategory', ''],
    ['hdnProcurementID', ''],
    ['hdnType', query.type || 'current'],
    ['hdnPreviousPge', 'TenderDetailsHome.html'],
    ['hdnadvsearch', ''],
    ['hdnFromStatus', ''],
    ['typeOfWorkFromConsolidation', ''],
    ['popUPRequestParameter', ''],
    ['selectedCircleDivison', ''],
    ['selectedDepartmentID', '0'],
    ['selectedProcurementType', ''],
    ['selectedTypeofWork', ''],
    ['aid', ''],
    ['hdnEncryptNames', 'hdnEncryptNames'],
    ['hdnEncryptValues', 'hdnEncryptValues'],
    ['sEcho', '1'],
    ['iColumns', '10'],
    ['sColumns', ',,,,,,,,,']
  ]
  params.push(['iDisplayStart', String(query.start)])
  params.push(['iDisplayLength', String(query.length)])
  for (let i = 0; i < 10; i += 1) {
    params.push([`mDataProp_${i}`, String(i)])
    params.push([`bSortable_${i}`, i === 9 ? 'false' : 'true'])
  }
  params.push(['iSortCol_0', '5'])
  params.push(['sSortDir_0', 'desc'])
  params.push(['iSortingCols', '1'])
  params.push(['_', String(Date.now())])
  const p = new URLSearchParams()
  for (const [k, v] of params) p.append(k, v)
  return `${TENDER_ORIGIN}/TenderDetailsHomeJson.html?${p.toString()}`
}

async function runTenderFetch(win: BrowserWindow, url: string): Promise<string> {
  return win.webContents.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {
       headers: {
         'x-requested-with': 'XMLHttpRequest',
         'accept': 'application/json, text/javascript, */*; q=0.01'
       },
       credentials: 'include'
     }).then(function (r) { return r.text() }).catch(function () { return '' })`
  )
}

const looksLikeJson = (s: string): boolean => s.trimStart().startsWith('{')

async function fetchTenders(query: TenderQuery): Promise<string> {
  const url = buildTenderUrl(query)

  // Fast path: reuse an already-seated window if one is alive.
  if (tenderWindow && !tenderWindow.isDestroyed()) {
    const body = await runTenderFetch(tenderWindow, url)
    if (looksLikeJson(body)) return body
  }

  // Otherwise (or on a stale session): rebuild a fresh window, seat it, retry.
  let body = ''
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const win = recreateTenderWindow()
    await seatTenderSession(win)
    body = await runTenderFetch(win, url)
    if (looksLikeJson(body)) return body
  }
  return body
}


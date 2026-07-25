import { app, BrowserWindow, clipboard, dialog, ipcMain, shell } from 'electron'
import * as fs from 'fs'
import * as path from 'path'
import * as https from 'https'
import * as os from 'os'
import * as firebaseSync from './firebaseSync'
import { initAutoUpdate, restartToUpdate, checkForUpdatesManually } from './autoUpdate'
import { IPC } from './ipc-contract'
import type { ManualCheckResult } from './ipc-contract'
import { parseExcelFile, readExcelGrid, readAllSheetGrids, buildWorkbookBuffer } from '../core/excel'
import { recognizeImage } from './ocr'
import { applyTechnicalSanctionEdits } from '../core/technicalSanctionOutput'
import type { CellEdit } from '../core/technicalSanction'
import { embedTexts } from './embeddings'
import { buildScheduleAWorkbook, rowsToScheduleAItems } from '../core/scheduleA'
import type { ScheduleAMeta } from '../core/scheduleA'
import { fillScheduleATemplate } from '../core/scheduleATemplate'
import { fillBoqTemplate, rowsToBoqData } from '../core/boqTemplate'
import { fillDeviationTemplate } from '../core/deviationTemplate'
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
import { fillTenderNotice } from '../core/tenderNotice'
import type { TenderNoticeInput } from '../core/tenderNotice'
import { fillBidDocument } from '../core/bidDocument'
import type { BidDocumentInput } from '../core/bidDocument'
import type { CalendarData } from '../core/calendar'
import { convertHtmlToDocx } from '../core/htmlToDocx'
import { convertDocxToPdf } from '../core/docxToPdf'
import { convertRtfToDocx } from '../core/rtfToDocx'
import {
  listParagraphs,
  applyParagraphEdits,
  findPlaceholdersInDocx,
  fillPlaceholdersInDocx,
  bakeFixedPlaceholdersInDocx
} from '../core/docx-edit'
import type { PlaceholderMatch } from '../core/createDocument'
import type {
  ExcelTable,
  TenderQuery,
  TenderResult,
  PersistedState
} from '../core/types'
import type { SheetGrid } from '../core/sheet'

let mainWindow: BrowserWindow | null = null

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 820,
    minWidth: 900,
    minHeight: 640,
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
      nodeIntegration: false
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
    return result.filePaths.map((p) => readExcelGrid(p))
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

  ipcMain.handle(IPC.ocrEstimatePhotos, async (_e, dataUrls: string[]): Promise<SheetGrid> => {
    const allRows: string[][] = []
    for (const dataUrl of dataUrls) {
      const base64 = dataUrl.split(',')[1] ?? ''
      const buffer = Buffer.from(base64, 'base64')
      const lines = await recognizeImage(buffer)
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

  // Scrape the Telangana Government 2026 holiday calendar page.
  ipcMain.handle(IPC.fetchCalendar, async (_e, force?: boolean) => {
    const cacheFile = path.join(app.getPath('userData'), 'calendar-cache.json')

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

    const url = 'https://www.telangana.gov.in/downloads/calendar-2026/'
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
    const parsed = parseCalendarHtml(html)
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

  ipcMain.handle(IPC.login, async (_e, loginId: string, password: string): Promise<LoginResult> => {
    const result = await validateLogin(loginId, password)
    if (!result.ok) return result

    const { claim, remoteState } = await firebaseSync.startSession(loginId, os.hostname(), (partial) => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send(IPC.remoteStateUpdate, partial)
      }
    })
    if (!claim.ok) return { ok: false, maxSessions: true }

    if (remoteState) {
      try {
        fs.writeFileSync(stateFile(), JSON.stringify(remoteState), 'utf8')
      } catch {
        // Non-fatal: local cache write is best-effort.
      }
    }
    return result
  })

  ipcMain.handle(IPC.logout, async (): Promise<void> => {
    await firebaseSync.endSession()
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

  ipcMain.handle(IPC.createDocumentFromClipboard, async (): Promise<string | null> => {
    // RTF is Word's own, far richer clipboard format — readBuffer (not the
    // string-returning readRTF) avoids corrupting any embedded image data
    // via a string encoding round-trip.
    const format = process.platform === 'darwin' ? 'public.rtf' : 'Rich Text Format'
    if (!clipboard.has(format)) return null
    const rtfBuffer = clipboard.readBuffer(format)
    if (rtfBuffer.length === 0) return null
    const docxBuffer = await convertRtfToDocx(rtfBuffer)
    return docxBuffer.toString('base64')
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
    return findPlaceholdersInDocx(Buffer.from(docxBase64, 'base64'))
  })

  ipcMain.handle(
    IPC.fillPlaceholdersInDocument,
    async (_e, docxBase64: string, resolved: PlaceholderMatch[], row: Record<string, string>): Promise<string> => {
      const filled = fillPlaceholdersInDocx(Buffer.from(docxBase64, 'base64'), resolved, row)
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

  ipcMain.handle(
    IPC.exportCreatedDocument,
    async (
      _e,
      docxBase64: string,
      suggestedName: string,
      formats: ('docx' | 'pdf')[]
    ): Promise<{ file: string; format: 'docx' | 'pdf' }[] | null> => {
      if (formats.length === 0) return null
      const docxBuffer = Buffer.from(docxBase64, 'base64')
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
    const body = /<html/i.test(renderedHtml) ? renderedHtml : `<!DOCTYPE html><html><body>${renderedHtml}</body></html>`
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
    win.webContents.print({}, (success, failureReason) => {
      if (!success && failureReason !== 'cancelled') {
        console.error('Print failed:', failureReason)
      }
    })
  })

  const stateFile = () => path.join(app.getPath('userData'), 'state.json')

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

      const filled = fillTenderNotice(fs.readFileSync(templatePath), input)
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
      const templatePath = bidDocumentTemplateFile()
      if (!templatePath) throw new Error('Bid Document template is missing from the app bundle.')

      const result = await dialog.showSaveDialog(mainWindow!, {
        title: 'Save Bid Document',
        defaultPath: `${suggestedName}.docx`,
        filters: [{ name: 'Word Document', extensions: ['docx'] }]
      })
      if (result.canceled || !result.filePath) return null

      const filled = fillBidDocument(fs.readFileSync(templatePath), input)
      fs.writeFileSync(result.filePath, filled)
      return result.filePath
    }
  )

  ipcMain.handle(
    IPC.previewBidDocument,
    async (_e, input: BidDocumentInput): Promise<string> => {
      const templatePath = bidDocumentTemplateFile()
      if (!templatePath) throw new Error('Bid Document template is missing from the app bundle.')
      const filled = fillBidDocument(fs.readFileSync(templatePath), input)
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
    return state && (await migrateCreatedDocuments(state))
  })

  ipcMain.handle(IPC.saveState, async (_e, state: PersistedState): Promise<void> => {
    try {
      fs.writeFileSync(stateFile(), JSON.stringify(state), 'utf8')
    } catch {
      // Non-fatal: persistence is best-effort.
    }
    // Don't block the local save on network latency — the cloud push is
    // itself best-effort and swallows its own errors.
    void firebaseSync.pushState(state)
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


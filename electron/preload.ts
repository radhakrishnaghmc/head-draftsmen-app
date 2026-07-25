import { contextBridge, ipcRenderer } from 'electron'
import { IPC } from './ipc-contract'
import type { DocuGenApi } from './ipc-contract'

const api: DocuGenApi = {
  pickExcels: () => ipcRenderer.invoke(IPC.pickExcels),
  pickExcelGrids: () => ipcRenderer.invoke(IPC.pickExcelGrids),
  pickEstimateGrid: () => ipcRenderer.invoke(IPC.pickEstimateGrid),
  pickDataSheet: () => ipcRenderer.invoke(IPC.pickDataSheet),
  ocrEstimatePhotos: (dataUrls) => ipcRenderer.invoke(IPC.ocrEstimatePhotos, dataUrls),
  openPath: (target) => ipcRenderer.invoke(IPC.openPath, target),
  revealItem: (target) => ipcRenderer.invoke(IPC.revealItem, target),
  defaultDir: () => ipcRenderer.invoke(IPC.defaultDir),
  getAppVersion: () => ipcRenderer.invoke(IPC.getAppVersion),
  fetchCalendar: (force?: boolean) => ipcRenderer.invoke(IPC.fetchCalendar, force),
  login: (loginId, password) => ipcRenderer.invoke(IPC.login, loginId, password),
  logout: () => ipcRenderer.invoke(IPC.logout),
  importFromLink: (url) => ipcRenderer.invoke(IPC.importFromLink, url),
  importAllSheetsFromLink: (url) => ipcRenderer.invoke(IPC.importAllSheetsFromLink, url),
  exportTable: (table, suggestedName) => ipcRenderer.invoke(IPC.exportTable, table, suggestedName),
  exportScheduleA: (table, suggestedName, meta) =>
    ipcRenderer.invoke(IPC.exportScheduleA, table, suggestedName, meta),
  exportBoq: (table, suggestedName, workName) =>
    ipcRenderer.invoke(IPC.exportBoq, table, suggestedName, workName),
  exportBoqBatch: (entries) => ipcRenderer.invoke(IPC.exportBoqBatch, entries),
  exportDeviation: (items, meta, suggestedName) =>
    ipcRenderer.invoke(IPC.exportDeviation, items, meta, suggestedName),
  exportDetailedEstimate: (items, meta, suggestedName) =>
    ipcRenderer.invoke(IPC.exportDetailedEstimate, items, meta, suggestedName),
  exportMaterialEstimate: (totals, meta, suggestedName) =>
    ipcRenderer.invoke(IPC.exportMaterialEstimate, totals, meta, suggestedName),
  generateTenderNotice: (input, suggestedName) =>
    ipcRenderer.invoke(IPC.generateTenderNotice, input, suggestedName),
  previewTenderNotice: (input) => ipcRenderer.invoke(IPC.previewTenderNotice, input),
  generateBidDocument: (input, suggestedName) =>
    ipcRenderer.invoke(IPC.generateBidDocument, input, suggestedName),
  generateBidDocumentBatch: (entries) => ipcRenderer.invoke(IPC.generateBidDocumentBatch, entries),
  previewBidDocument: (input) => ipcRenderer.invoke(IPC.previewBidDocument, input),
  generateTechnicalSanction: (estimatePath, sheetName, edits, suggestedName, rateAnalysisRows) =>
    ipcRenderer.invoke(
      IPC.generateTechnicalSanction,
      estimatePath,
      sheetName,
      edits,
      suggestedName,
      rateAnalysisRows
    ),
  searchTenders: (query) => ipcRenderer.invoke(IPC.searchTenders, query),
  embedTexts: (texts) => ipcRenderer.invoke(IPC.embedTexts, texts),
  createDocumentFromClipboard: () => ipcRenderer.invoke(IPC.createDocumentFromClipboard),
  listDocumentParagraphs: (docxBase64) => ipcRenderer.invoke(IPC.listDocumentParagraphs, docxBase64),
  saveDocumentEdits: (docxBase64, edits) => ipcRenderer.invoke(IPC.saveDocumentEdits, docxBase64, edits),
  findPlaceholdersInDocument: (docxBase64) => ipcRenderer.invoke(IPC.findPlaceholdersInDocument, docxBase64),
  fillPlaceholdersInDocument: (docxBase64, resolved, row) =>
    ipcRenderer.invoke(IPC.fillPlaceholdersInDocument, docxBase64, resolved, row),
  bakeFixedPlaceholdersInDocument: (docxBase64, values) =>
    ipcRenderer.invoke(IPC.bakeFixedPlaceholdersInDocument, docxBase64, values),
  exportCreatedDocument: (docxBase64, suggestedName, formats) =>
    ipcRenderer.invoke(IPC.exportCreatedDocument, docxBase64, suggestedName, formats),
  printCreatedDocument: (renderedHtml) => ipcRenderer.invoke(IPC.printCreatedDocument, renderedHtml),
  exportIntimationHtml: (html, suggestedName) => ipcRenderer.invoke(IPC.exportIntimationHtml, html, suggestedName),
  loadState: () => ipcRenderer.invoke(IPC.loadState),
  saveState: (state) => ipcRenderer.invoke(IPC.saveState, state),
  onRemoteStateUpdate: (callback) => {
    const listener = (_e: unknown, partial: import('../core/types').PersistedState) => callback(partial)
    ipcRenderer.on(IPC.remoteStateUpdate, listener)
    return () => ipcRenderer.removeListener(IPC.remoteStateUpdate, listener)
  },
  onUpdateDownloaded: (callback) => {
    const listener = () => callback()
    ipcRenderer.on(IPC.updateDownloaded, listener)
    return () => ipcRenderer.removeListener(IPC.updateDownloaded, listener)
  },
  onUpdateProgress: (callback) => {
    const listener = (_e: unknown, progress: import('./ipc-contract').UpdateProgress) => callback(progress)
    ipcRenderer.on(IPC.updateProgress, listener)
    return () => ipcRenderer.removeListener(IPC.updateProgress, listener)
  },
  restartToUpdate: () => ipcRenderer.send(IPC.restartToUpdate),
  onUpdateInstallError: (callback) => {
    const listener = (_e: unknown, message: string) => callback(message)
    ipcRenderer.on(IPC.updateInstallError, listener)
    return () => ipcRenderer.removeListener(IPC.updateInstallError, listener)
  },
  checkForUpdates: () => ipcRenderer.invoke(IPC.checkForUpdates)
}

contextBridge.exposeInMainWorld('docugen', api)

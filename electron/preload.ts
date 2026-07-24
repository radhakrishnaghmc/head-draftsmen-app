import { contextBridge, ipcRenderer } from 'electron'
import { IPC } from './ipc-contract'
import type { DocuGenApi } from './ipc-contract'

const api: DocuGenApi = {
  pickExcels: () => ipcRenderer.invoke(IPC.pickExcels),
  pickExcelGrids: () => ipcRenderer.invoke(IPC.pickExcelGrids),
  pickEstimateGrid: () => ipcRenderer.invoke(IPC.pickEstimateGrid),
  pickDataSheet: () => ipcRenderer.invoke(IPC.pickDataSheet),
  openPath: (target) => ipcRenderer.invoke(IPC.openPath, target),
  revealItem: (target) => ipcRenderer.invoke(IPC.revealItem, target),
  defaultDir: () => ipcRenderer.invoke(IPC.defaultDir),
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
  generateTenderNotice: (input, suggestedName) =>
    ipcRenderer.invoke(IPC.generateTenderNotice, input, suggestedName),
  previewTenderNotice: (input) => ipcRenderer.invoke(IPC.previewTenderNotice, input),
  generateBidDocument: (input, suggestedName) =>
    ipcRenderer.invoke(IPC.generateBidDocument, input, suggestedName),
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
  exportCreatedDocument: (html, suggestedName, formats) =>
    ipcRenderer.invoke(IPC.exportCreatedDocument, html, suggestedName, formats),
  printCreatedDocument: (html) => ipcRenderer.invoke(IPC.printCreatedDocument, html),
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
  restartToUpdate: () => ipcRenderer.send(IPC.restartToUpdate)
}

contextBridge.exposeInMainWorld('docugen', api)

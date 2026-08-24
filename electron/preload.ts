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
  fetchCalendar: (url, force) => ipcRenderer.invoke(IPC.fetchCalendar, url, force),
  login: (loginId, password) => ipcRenderer.invoke(IPC.login, loginId, password),
  logout: () => ipcRenderer.invoke(IPC.logout),
  importFromLink: (url) => ipcRenderer.invoke(IPC.importFromLink, url),
  importAllSheetsFromLink: (url) => ipcRenderer.invoke(IPC.importAllSheetsFromLink, url),
  exportTable: (table, suggestedName) => ipcRenderer.invoke(IPC.exportTable, table, suggestedName),
  exportEvaluationSheet: (input, suggestedName) => ipcRenderer.invoke(IPC.exportEvaluationSheet, input, suggestedName),
  exportScheduleA: (table, suggestedName, meta) =>
    ipcRenderer.invoke(IPC.exportScheduleA, table, suggestedName, meta),
  exportAgreementBundle: (files) => ipcRenderer.invoke(IPC.exportAgreementBundle, files),
  exportBoq: (table, suggestedName, workName) =>
    ipcRenderer.invoke(IPC.exportBoq, table, suggestedName, workName),
  exportBoqBatch: (entries) => ipcRenderer.invoke(IPC.exportBoqBatch, entries),
  pickWorkbookForSplit: () => ipcRenderer.invoke(IPC.pickWorkbookForSplit),
  splitWorkbook: (srcPath, sheetNames) => ipcRenderer.invoke(IPC.splitWorkbook, srcPath, sheetNames),
  pickPdfsForMerge: () => ipcRenderer.invoke(IPC.pickPdfsForMerge),
  mergePdfs: (srcPaths) => ipcRenderer.invoke(IPC.mergePdfs, srcPaths),
  pickPdfForSplit: () => ipcRenderer.invoke(IPC.pickPdfForSplit),
  splitPdf: (srcPath, ranges) => ipcRenderer.invoke(IPC.splitPdf, srcPath, ranges),
  savePdf: (bytes, suggestedName) => ipcRenderer.invoke(IPC.savePdf, bytes, suggestedName),
  savePdfsToFolder: (files) => ipcRenderer.invoke(IPC.savePdfsToFolder, files),
  docxToPdf: (docxBytes) => ipcRenderer.invoke(IPC.docxToPdf, docxBytes),
  docxToPageImages: (docxBytes) => ipcRenderer.invoke(IPC.docxToPageImages, docxBytes),
  mergeDocx: (docxBytesList) => ipcRenderer.invoke(IPC.mergeDocx, docxBytesList),
  splitDocxSections: (docxBytes) => ipcRenderer.invoke(IPC.splitDocxSections, docxBytes),
  saveDocxsToFolder: (files) => ipcRenderer.invoke(IPC.saveDocxsToFolder, files),
  ocrGpsOverlay: (imageBytes) => ipcRenderer.invoke(IPC.ocrGpsOverlay, imageBytes),
  ocrPhotosToLines: (dataUrls) => ipcRenderer.invoke(IPC.ocrPhotosToLines, dataUrls),
  savePhotosAsWord: (text, suggestedName) => ipcRenderer.invoke(IPC.savePhotosAsWord, text, suggestedName),
  savePhotosAsExcel: (text, suggestedName) => ipcRenderer.invoke(IPC.savePhotosAsExcel, text, suggestedName),
  savePdfAsWord: (pdfs, suggestedName) => ipcRenderer.invoke(IPC.savePdfAsWord, pdfs, suggestedName),
  saveWordDoc: (blocks, suggestedName) => ipcRenderer.invoke(IPC.saveWordDoc, blocks, suggestedName),
  ocrPhotosToLayout: (dataUrls) => ipcRenderer.invoke(IPC.ocrPhotosToLayout, dataUrls),
  saveRowsAsExcel: (rows, suggestedName) => ipcRenderer.invoke(IPC.saveRowsAsExcel, rows, suggestedName),
  onSplitProgress: (callback) => {
    const listener = (_e: unknown, progress: import('./ipc-contract').SplitProgress) => callback(progress)
    ipcRenderer.on(IPC.splitProgress, listener)
    return () => ipcRenderer.removeListener(IPC.splitProgress, listener)
  },
  onAgreementBundleProgress: (callback) => {
    const listener = (_e: unknown, progress: import('./ipc-contract').AgreementBundleProgress) => callback(progress)
    ipcRenderer.on(IPC.agreementBundleProgress, listener)
    return () => ipcRenderer.removeListener(IPC.agreementBundleProgress, listener)
  },
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
  noteSubmittedDocx: (html) => ipcRenderer.invoke(IPC.noteSubmittedDocx, html),
  intimationTemplate: () => ipcRenderer.invoke(IPC.intimationTemplate),
  workOrderTemplate: (variantId) => ipcRenderer.invoke(IPC.workOrderTemplate, variantId),
  fileBackerTemplate: () => ipcRenderer.invoke(IPC.fileBackerTemplate),
  agreementTemplate: (variantId) => ipcRenderer.invoke(IPC.agreementTemplate, variantId),
  qccIntimationTemplate: () => ipcRenderer.invoke(IPC.qccIntimationTemplate),
  forwardingSlipTemplate: () => ipcRenderer.invoke(IPC.forwardingSlipTemplate),
  civilTenderTemplate: () => ipcRenderer.invoke(IPC.civilTenderTemplate),
  seAgreementBondTemplate: () => ipcRenderer.invoke(IPC.seAgreementBondTemplate),
  zonalWorkOrderTemplate: () => ipcRenderer.invoke(IPC.zonalWorkOrderTemplate),
  zonalConcludingAgreementTemplate: () => ipcRenderer.invoke(IPC.zonalConcludingAgreementTemplate),
  zonalMemoEeTemplate: () => ipcRenderer.invoke(IPC.zonalMemoEeTemplate),
  seAgreementNoteTemplate: () => ipcRenderer.invoke(IPC.seAgreementNoteTemplate),
  contractDeedTemplate: () => ipcRenderer.invoke(IPC.contractDeedTemplate),
  exportSeScheduleA: (table, suggestedName, meta) => ipcRenderer.invoke(IPC.exportSeScheduleA, table, suggestedName, meta),
  loaSeTemplate: (reserved) => ipcRenderer.invoke(IPC.loaSeTemplate, reserved),
  tsNoteTemplate: () => ipcRenderer.invoke(IPC.tsNoteTemplate),
  eligibilityCriteriaTemplate: () => ipcRenderer.invoke(IPC.eligibilityCriteriaTemplate),
  issueNoticeTemplate: (ee) => ipcRenderer.invoke(IPC.issueNoticeTemplate, ee),
  loadState: () => ipcRenderer.invoke(IPC.loadState),
  saveState: (state, skipCloud) => ipcRenderer.invoke(IPC.saveState, state, skipCloud),
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

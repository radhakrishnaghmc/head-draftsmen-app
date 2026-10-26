import { contextBridge, ipcRenderer } from 'electron'
import { IPC } from './ipc-contract'
import type { DocuGenApi } from './ipc-contract'

const api: DocuGenApi = {
  pickExcels: () => ipcRenderer.invoke(IPC.pickExcels),
  pickExcelGrids: () => ipcRenderer.invoke(IPC.pickExcelGrids),
  pickExcelPaths: () => ipcRenderer.invoke(IPC.pickExcelPaths),
  readWorkbookGrids: (filePath: string) => ipcRenderer.invoke(IPC.readWorkbookGrids, filePath),
  pickEstimateGrid: () => ipcRenderer.invoke(IPC.pickEstimateGrid),
  pickEstimateGrids: () => ipcRenderer.invoke(IPC.pickEstimateGrids),
  pickEstimateGridsMulti: () => ipcRenderer.invoke(IPC.pickEstimateGridsMulti),
  pickDataSheet: () => ipcRenderer.invoke(IPC.pickDataSheet),
  pickTenderDocuments: () => ipcRenderer.invoke(IPC.pickTenderDocuments),
  onTenderScanProgress: (callback) => {
    const listener = (_e: unknown, progress: import('./ipc-contract').TenderScanProgress) => callback(progress)
    ipcRenderer.on(IPC.tenderScanProgress, listener)
    return () => ipcRenderer.removeListener(IPC.tenderScanProgress, listener)
  },
  ocrEstimatePhotos: (dataUrls) => ipcRenderer.invoke(IPC.ocrEstimatePhotos, dataUrls),
  ocrMbMeasurementSheet: (dataUrls) => ipcRenderer.invoke(IPC.ocrMbMeasurementSheet, dataUrls),
  onMbMeasurementProgress: (callback) => {
    const listener = (_e: unknown, progress: import('./ipc-contract').MbMeasurementProgress) => callback(progress)
    ipcRenderer.on(IPC.mbMeasurementProgress, listener)
    return () => ipcRenderer.removeListener(IPC.mbMeasurementProgress, listener)
  },
  openPath: (target) => ipcRenderer.invoke(IPC.openPath, target),
  revealItem: (target) => ipcRenderer.invoke(IPC.revealItem, target),
  defaultDir: () => ipcRenderer.invoke(IPC.defaultDir),
  getAppVersion: () => ipcRenderer.invoke(IPC.getAppVersion),
  fetchCalendar: (url, force) => ipcRenderer.invoke(IPC.fetchCalendar, url, force),
  login: (loginId, password, forceLogout) => ipcRenderer.invoke(IPC.login, loginId, password, forceLogout),
  logout: () => ipcRenderer.invoke(IPC.logout),
  listActiveSessions: () => ipcRenderer.invoke(IPC.listActiveSessions),
  logoutOtherSession: (sessionId) => ipcRenderer.invoke(IPC.logoutOtherSession, sessionId),
  listShareUsers: (corporation) => ipcRenderer.invoke(IPC.listShareUsers, corporation),
  shareTenderConditions: (corporation, targets, entries) => ipcRenderer.invoke(IPC.shareTenderConditions, corporation, targets, entries),
  importFromLink: (url) => ipcRenderer.invoke(IPC.importFromLink, url),
  importAllSheetsFromLink: (url) => ipcRenderer.invoke(IPC.importAllSheetsFromLink, url),
  exportTable: (table, suggestedName) => ipcRenderer.invoke(IPC.exportTable, table, suggestedName),
  exportEvaluationSheet: (input, suggestedName) => ipcRenderer.invoke(IPC.exportEvaluationSheet, input, suggestedName),
  exportEvaluationSheetBatch: (entries) => ipcRenderer.invoke(IPC.exportEvaluationSheetBatch, entries),
  exportScheduleA: (table, suggestedName, meta) =>
    ipcRenderer.invoke(IPC.exportScheduleA, table, suggestedName, meta),
  scheduleAPreviewHtml: (table, meta, isSe) => ipcRenderer.invoke(IPC.scheduleAPreviewHtml, table, meta, isSe),
  exportScheduleAPdf: (table, suggestedName, meta, isSe) =>
    ipcRenderer.invoke(IPC.exportScheduleAPdf, table, suggestedName, meta, isSe),
  exportAgreementBundle: (files, mergedPdfName) => ipcRenderer.invoke(IPC.exportAgreementBundle, files, mergedPdfName),
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
  fetchWeather: (apiKey, city) => ipcRenderer.invoke(IPC.fetchWeather, apiKey, city),
  fetchCementSteelRates: () => ipcRenderer.invoke(IPC.fetchCementSteelRates),
  downloadCementSteelRate: (token, suggestedFileName) =>
    ipcRenderer.invoke(IPC.downloadCementSteelRate, token, suggestedFileName),
  fetchCementSteelRateBytes: (token) => ipcRenderer.invoke(IPC.fetchCementSteelRateBytes, token),
  ocrGpsOverlay: (imageBytes) => ipcRenderer.invoke(IPC.ocrGpsOverlay, imageBytes),
  ocrPhotosToLines: (dataUrls) => ipcRenderer.invoke(IPC.ocrPhotosToLines, dataUrls),
  savePhotosAsWord: (text, suggestedName) => ipcRenderer.invoke(IPC.savePhotosAsWord, text, suggestedName),
  savePhotosAsExcel: (text, suggestedName) => ipcRenderer.invoke(IPC.savePhotosAsExcel, text, suggestedName),
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
  onBidDocumentBatchProgress: (callback) => {
    const listener = (_e: unknown, progress: import('./ipc-contract').BidDocumentBatchProgress) => callback(progress)
    ipcRenderer.on(IPC.bidDocumentBatchProgress, listener)
    return () => ipcRenderer.removeListener(IPC.bidDocumentBatchProgress, listener)
  },
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
  fetchTenderConditions: (tenderId) => ipcRenderer.invoke(IPC.fetchTenderConditions, tenderId),
  viewTenderDetails: (indentId, category, procurementId) =>
    ipcRenderer.invoke(IPC.viewTenderDetails, indentId, category, procurementId),
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
  printCreatedDocument: (renderedHtml, pageMarginsMm) =>
    ipcRenderer.invoke(IPC.printCreatedDocument, renderedHtml, pageMarginsMm),
  fontFallbackFiles: () => ipcRenderer.invoke(IPC.fontFallbackFiles),
  noteSubmittedDocx: (html, marginsTwips) => ipcRenderer.invoke(IPC.noteSubmittedDocx, html, marginsTwips),
  intimationTemplate: (variantId) => ipcRenderer.invoke(IPC.intimationTemplate, variantId),
  workOrderTemplate: (variantId) => ipcRenderer.invoke(IPC.workOrderTemplate, variantId),
  fileBackerTemplate: (variantId) => ipcRenderer.invoke(IPC.fileBackerTemplate, variantId),
  agreementTemplate: (variantId) => ipcRenderer.invoke(IPC.agreementTemplate, variantId),
  qccIntimationTemplate: (variantId) => ipcRenderer.invoke(IPC.qccIntimationTemplate, variantId),
  forwardingSlipTemplate: (variantId?: string) => ipcRenderer.invoke(IPC.forwardingSlipTemplate, variantId),
  workBillNoteTemplate: () => ipcRenderer.invoke(IPC.workBillNoteTemplate),
  aeDyeeForwardingTemplate: (variantId) => ipcRenderer.invoke(IPC.aeDyeeForwardingTemplate, variantId),
  civilTenderTemplate: (variantId) => ipcRenderer.invoke(IPC.civilTenderTemplate, variantId),
  seAgreementBondTemplate: () => ipcRenderer.invoke(IPC.seAgreementBondTemplate),
  zonalWorkOrderTemplate: (variantId) => ipcRenderer.invoke(IPC.zonalWorkOrderTemplate, variantId),
  zonalConcludingAgreementTemplate: (variantId) => ipcRenderer.invoke(IPC.zonalConcludingAgreementTemplate, variantId),
  zonalMemoEeTemplate: (variantId) => ipcRenderer.invoke(IPC.zonalMemoEeTemplate, variantId),
  seAgreementNoteTemplate: () => ipcRenderer.invoke(IPC.seAgreementNoteTemplate),
  contractDeedTemplate: () => ipcRenderer.invoke(IPC.contractDeedTemplate),
  exportSeScheduleA: (table, suggestedName, meta) => ipcRenderer.invoke(IPC.exportSeScheduleA, table, suggestedName, meta),
  loaSeTemplate: (reserved, corporation) => ipcRenderer.invoke(IPC.loaSeTemplate, reserved, corporation),
  verifyDocuments: (items) => ipcRenderer.invoke(IPC.verifyDocuments, items),
  tsNoteTemplate: () => ipcRenderer.invoke(IPC.tsNoteTemplate),
  eligibilityCriteriaTemplate: () => ipcRenderer.invoke(IPC.eligibilityCriteriaTemplate),
  issueNoticeTemplate: (ee, variantId) => ipcRenderer.invoke(IPC.issueNoticeTemplate, ee, variantId),
  loadState: () => ipcRenderer.invoke(IPC.loadState),
  saveState: (state, skipCloud) => ipcRenderer.invoke(IPC.saveState, state, skipCloud),
  syncNow: (state) => ipcRenderer.invoke(IPC.syncNow, state),
  createBackup: (state) => ipcRenderer.invoke(IPC.createBackup, state),
  revealBackup: (filePath) => ipcRenderer.invoke(IPC.revealBackup, filePath),
  saveBackupCopy: (filePath) => ipcRenderer.invoke(IPC.saveBackupCopy, filePath),
  restoreBackup: (currentState) => ipcRenderer.invoke(IPC.restoreBackup, currentState),
  getActivityLog: () => ipcRenderer.invoke(IPC.getActivityLog),
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
